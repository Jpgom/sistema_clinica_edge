"""Snapshots consistentes e restauração que conserva efeitos SMTP já realizados."""
from contextlib import closing
from datetime import datetime
from pathlib import Path
import hashlib
import json
import re
import shutil
import sqlite3
import uuid
import zipfile

from send_queue import job_lock


def _digest(path):
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def _asset_path(api, name, control=False):
    name = str(name or "").replace("\\", "/")
    relative = Path(name)
    if control and relative.parts and relative.parts[0] != "controls":
        relative = Path("controls") / relative
    result = (api.DATA_DIR / relative).resolve()
    if not result.is_relative_to(api.DATA_DIR.resolve()) or not name:
        raise ValueError("O banco contém uma referência de arquivo inválida.")
    return result


def _snapshot_references(api, db_path):
    with closing(sqlite3.connect(db_path)) as conn:
        conn.row_factory = sqlite3.Row
        docs = conn.execute("SELECT stored_name,sha256 FROM documents ORDER BY id").fetchall()
        controls = conn.execute("SELECT control_stored_name,control_hash FROM competencies WHERE control_stored_name IS NOT NULL AND control_stored_name!=''").fetchall()
    refs = {}
    for row in docs:
        path = _asset_path(api, row["stored_name"])
        refs[path] = row["sha256"]
    for row in controls:
        path = _asset_path(api, row["control_stored_name"], control=True)
        refs[path] = row["control_hash"]
    return refs


def create(api, label="backup", allow_missing_assets=False):
    """Só publica .db depois de terminar e validar seu pacote de arquivos."""
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", str(label or "backup"))[:60] or "backup"
    target = api.BACKUP_DIR / f"{datetime.now():%Y%m%d_%H%M%S_%f}_{safe}.db"
    staging = api.DATA_DIR / (".backup_" + uuid.uuid4().hex)
    staging.mkdir()
    staged_db, staged_zip = staging / "snapshot.db", staging / "snapshot.zip"
    published = []
    try:
        with closing(api.db()) as src, closing(sqlite3.connect(staged_db)) as dst:
            src.backup(dst)
        references = _snapshot_references(api, staged_db)
        missing, included = [], set()
        with zipfile.ZipFile(staged_zip, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(staged_db, "cobrancas.db")
            for path, expected_hash in references.items():
                relative = path.relative_to(api.DATA_DIR.resolve()).as_posix()
                if not path.is_file():
                    if allow_missing_assets:
                        missing.append(relative)
                        continue
                    raise ValueError(f"Backup não criado: arquivo referenciado ausente ({relative}). Reimporte ou restaure o arquivo.")
                contents = path.read_bytes()
                if expected_hash and hashlib.sha256(contents).hexdigest() != expected_hash:
                    if not allow_missing_assets:
                        raise ValueError(f"Backup não criado: arquivo referenciado foi alterado ({relative}). Confira o documento original.")
                    missing.append(relative + " (conteúdo alterado)")
                archive.writestr(relative, contents)
                included.add(relative)
            for folder in (api.DOCS_DIR, api.CONTROL_DIR):
                for path in sorted(folder.rglob("*")):
                    relative = path.relative_to(api.DATA_DIR).as_posix()
                    if path.is_file() and relative not in included:
                        try:
                            archive.write(path, relative)
                        except FileNotFoundError:
                            continue
                        included.add(relative)
            for name in (".flask_secret", ".fernet_key"):
                path = api.DATA_DIR / name
                if path.is_file():
                    archive.write(path, name)
            archive.writestr("backup_manifest.json", json.dumps({"format": "EDGE_BACKUP_V1", "missing_or_changed_files": missing}, ensure_ascii=False))
        staged_zip.replace(target.with_suffix(".zip"))
        published.append(target.with_suffix(".zip"))
        staged_db.replace(target)
        published.append(target)
        for old in sorted(api.BACKUP_DIR.glob("*.db"), key=lambda p: p.stat().st_mtime, reverse=True)[30:]:
            old.unlink(missing_ok=True)
            old.with_suffix(".zip").unlink(missing_ok=True)
        return target
    except Exception:
        for path in published:
            path.unlink(missing_ok=True)
        raise
    finally:
        if staging.resolve().is_relative_to(api.DATA_DIR.resolve()):
            shutil.rmtree(staging)


def _identity(row):
    return (str(row["unit_name"]).strip().casefold(), int(row["month"]), int(row["year"]), re.sub(r"\D", "", row["cnpj"]))


def _capture_external_knowledge(api, staging):
    """Entrega SMTP não pode ser desfeita por rollback do banco local."""
    records, logs, docs = {}, [], []
    with closing(api.db()) as conn:
        rows = conn.execute("""SELECT cc.*,c.cnpj,c.name company_name,c.email,c.email_cc,c.bill_complementaries,
            u.name unit_name,cp.month,cp.year FROM competency_companies cc
            JOIN companies c ON c.id=cc.company_id JOIN competencies cp ON cp.id=cc.competency_id
            JOIN units u ON u.id=cp.unit_id""").fetchall()
        by_pair = {(r["competency_id"], r["company_id"]): dict(r) for r in rows}
        for row in rows:
            if row["fixed_sent_at"] or row["complementary_sent_at"]:
                records[_identity(row)] = dict(row)
        for log in conn.execute("SELECT * FROM email_logs WHERE status IN ('ENVIADO','ENVIO_PARCIAL') ORDER BY id"):
            row = by_pair.get((log["competency_id"], log["company_id"]))
            if not row:
                continue
            key = _identity(row)
            field = "fixed" if log["email_type"].startswith("FIXED") else "complementary"
            known = records.setdefault(key, dict(row))
            if not known[field + "_sent_at"]:
                known[field + "_sent_at"] = log["sent_at"] or log["created_at"]
                known[field + "_amount"] = log["amount"] if log["amount"] is not None else row[field + "_amount"]
            logs.append((key, dict(log)))
        doc_requests = {(key, billing) for key, record in records.items()
                        for billing, field in (("FIXED", "fixed"), ("COMPLEMENTARY", "complementary")) if record[field + "_sent_at"]}
        # A responsável do grupo pode não ter cobrança própria. O recibo do
        # grupo identifica seus anexos mesmo quando o carimbo está só nos membros.
        for group in conn.execute("""SELECT j.competency_id,j.email_type,g.company_id,g.status,g.delivery_result_json
            FROM send_job_groups g JOIN send_jobs j ON j.id=g.job_id
            WHERE j.test_mode=0 AND g.status IN ('SENT','PARTIAL')"""):
            result = json.loads(group["delivery_result_json"] or "{}")
            if group["status"] == "PARTIAL" and not result.get("primary_accepted"):
                continue
            owner = by_pair.get((group["competency_id"], group["company_id"]))
            if owner:
                key = _identity(owner)
                records.setdefault(key, dict(owner))
                doc_requests.add((key, group["email_type"].split("_", 1)[0]))
        missing = 0
        for key, billing in sorted(doc_requests):
            record = records[key]
            for doc in conn.execute("SELECT * FROM documents WHERE competency_id=? AND company_id=? AND billing_type=? ORDER BY id",
                                    (record["competency_id"], record["company_id"], billing)):
                path = _asset_path(api, doc["stored_name"])
                if not path.is_file():
                    missing += 1
                    continue
                staged = staging / "preserved" / (uuid.uuid4().hex + path.suffix)
                staged.parent.mkdir(exist_ok=True)
                shutil.copy2(path, staged)
                docs.append((key, dict(doc), staged))
    return records, logs, docs, missing


def _validate_source(api, staged_db, staged_zip, staging):
    with closing(sqlite3.connect(staged_db)) as check:
        integrity = check.execute("PRAGMA integrity_check").fetchone()[0]
        tables = {r[0] for r in check.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if integrity != "ok" or not {"companies", "competencies", "settings", "competency_companies"}.issubset(tables):
            raise ValueError("Backup inválido ou corrompido.")
        if check.execute("PRAGMA foreign_key_check").fetchone():
            raise ValueError("Backup com referências de banco inválidas.")
        reference_names = {str(r[0]).replace("\\", "/") for r in check.execute("SELECT stored_name FROM documents")} if "documents" in tables else set()
    assets, keys = [], {}
    if staged_zip.is_file():
        with zipfile.ZipFile(staged_zip) as archive:
            if archive.testzip() is not None:
                raise ValueError("O pacote de arquivos do backup está corrompido.")
            if "cobrancas.db" not in archive.namelist():
                raise ValueError("O pacote não contém o banco correspondente ao backup.")
            with archive.open("cobrancas.db") as bundled_db:
                if hashlib.file_digest(bundled_db, "sha256").hexdigest() != _digest(staged_db):
                    raise ValueError("O arquivo .db e seu pacote ZIP pertencem a backups diferentes.")
            manifest = json.loads(archive.read("backup_manifest.json")) if "backup_manifest.json" in archive.namelist() else {}
            known_missing = set(manifest.get("missing_or_changed_files", []))
            for member in archive.infolist():
                if member.is_dir() or member.filename in {"cobrancas.db", "backup_manifest.json"}:
                    continue
                name = Path(member.filename.replace("\\", "/"))
                if not name.parts or not (name.parts[0] in {"documents", "controls"} or member.filename in {".flask_secret", ".fernet_key"} or name.as_posix() in reference_names):
                    raise ValueError("O pacote de backup contém um arquivo inesperado.")
                destination = (api.DATA_DIR / name).resolve()
                staged_asset = (staging / "assets" / name).resolve()
                if not destination.is_relative_to(api.DATA_DIR.resolve()) or not staged_asset.is_relative_to(staging.resolve()):
                    raise ValueError("Caminho inválido no pacote de backup.")
                staged_asset.parent.mkdir(parents=True, exist_ok=True)
                staged_asset.write_bytes(archive.read(member))
                assets.append((staged_asset, destination))
                if member.filename in {".flask_secret", ".fernet_key"}:
                    keys[member.filename] = staged_asset
        staged_files = {destination: staged for staged, destination in assets}
        unavailable = []
        if "documents" in tables:
            for destination, expected_hash in _snapshot_references(api, staged_db).items():
                name = destination.relative_to(api.DATA_DIR.resolve()).as_posix()
                staged = staged_files.get(destination)
                if not staged:
                    if name in known_missing:
                        unavailable.append(name)
                        continue
                    raise ValueError(f"O pacote não contém um arquivo referenciado pelo banco: {name}")
                if expected_hash and _digest(staged) != expected_hash:
                    if name + " (conteúdo alterado)" in known_missing:
                        unavailable.append(name)
                        continue
                    raise ValueError(f"O conteúdo do arquivo não corresponde ao backup: {name}")
        keys["_missing_assets"] = unavailable
        if ".fernet_key" in keys:
            api.Fernet(keys[".fernet_key"].read_bytes().strip())
        if ".flask_secret" in keys:
            if not keys[".flask_secret"].read_text(encoding="utf-8").strip():
                raise ValueError("O pacote contém uma chave de sessão vazia.")
    return assets, keys


def _merge_external_knowledge(api, records, logs, docs, install_asset):
    now = api.now_iso()
    mapping = {}
    with closing(api.db()) as conn, conn:
        units = {str(row["name"]).strip().casefold(): row["id"] for row in conn.execute("SELECT id,name FROM units")}
        companies = {(row["unit_id"], re.sub(r"\D", "", row["cnpj"])): row["id"] for row in conn.execute("SELECT id,unit_id,cnpj FROM companies")}
        competencies = {(row["unit_id"], row["month"], row["year"]): row["id"] for row in conn.execute("SELECT id,unit_id,month,year FROM competencies")}
        for key, record in records.items():
            uid = units.get(key[0])
            if uid is None:
                uid = conn.execute("INSERT INTO units(name,active,created_at,updated_at) VALUES(?,1,?,?)", (record["unit_name"], now, now)).lastrowid
                units[key[0]] = uid
            cid = companies.get((uid, key[3]))
            if cid is None:
                cid = conn.execute(
                    "INSERT INTO companies(unit_id,cnpj,name,email,email_cc,fixed_value,bill_complementaries,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?)",
                    (uid, key[3], record["company_name"], record["email"], record["email_cc"], record["fixed_amount"], record["bill_complementaries"], now, now)).lastrowid
                companies[uid, key[3]] = cid
            comp_id = competencies.get((uid, record["month"], record["year"]))
            if comp_id is None:
                comp_id = conn.execute("INSERT INTO competencies(unit_id,month,year,status,created_at,updated_at) VALUES(?,?,?,'ENVIADA',?,?)",
                                       (uid, record["month"], record["year"], now, now)).lastrowid
                competencies[uid, record["month"], record["year"]] = comp_id
            conn.execute("INSERT OR IGNORE INTO competency_companies(competency_id,company_id,created_at,updated_at) VALUES(?,?,?,?)", (comp_id, cid, now, now))
            mapping[key] = (comp_id, cid)
            for field in ("fixed", "complementary"):
                if record[field + "_sent_at"]:
                    conn.execute(f"UPDATE competency_companies SET {field}_sent_at=?,{field}_amount=?,updated_at=? WHERE competency_id=? AND company_id=?",
                                 (record[field + "_sent_at"], record[field + "_amount"], now, comp_id, cid))
            conn.execute("UPDATE competencies SET status='ENVIADA',updated_at=? WHERE id=? AND COALESCE(status,'')!='FECHADA'", (now, comp_id))
        for key, log in logs:
            comp_id, cid = mapping[key]
            exists = conn.execute("""SELECT 1 FROM email_logs WHERE competency_id=? AND company_id=? AND email_type=?
                AND COALESCE(to_email,'')=? AND COALESCE(subject,'')=? AND COALESCE(sent_at,'')=? LIMIT 1""",
                (comp_id, cid, log["email_type"], log["to_email"] or "", log["subject"] or "", log["sent_at"] or "")).fetchone()
            if not exists:
                conn.execute("INSERT INTO email_logs(competency_id,company_id,email_type,to_email,subject,amount,status,error,sent_at,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                             (comp_id, cid, log["email_type"], log["to_email"], log["subject"], log["amount"], log["status"], log["error"], log["sent_at"], log["created_at"]))
        replaced_types = set()
        for key, doc, staged in docs:
            comp_id, cid = mapping[key]
            target_type = (comp_id, cid, doc["billing_type"])
            if target_type not in replaced_types:
                conn.execute("DELETE FROM documents WHERE competency_id=? AND company_id=? AND billing_type=?", target_type)
                replaced_types.add(target_type)
            sha = _digest(staged)
            if conn.execute("SELECT 1 FROM documents WHERE competency_id=? AND company_id=? AND billing_type=? AND sha256=?", (comp_id, cid, doc["billing_type"], sha)).fetchone():
                continue
            destination = api.DOCS_DIR / str(comp_id) / doc["billing_type"].lower() / str(cid) / ("restore_" + uuid.uuid4().hex + staged.suffix)
            install_asset(staged, destination)
            conn.execute("INSERT INTO documents(competency_id,company_id,billing_type,original_name,stored_name,sha256,created_at) VALUES(?,?,?,?,?,?,?)",
                         (comp_id, cid, doc["billing_type"], doc["original_name"], str(destination.relative_to(api.DATA_DIR)), sha, doc["created_at"]))
        jobs = conn.execute("SELECT id FROM send_jobs WHERE status IN ('QUEUED','RUNNING')").fetchall()
        for job in jobs:
            conn.execute("UPDATE send_job_groups SET status='UNCERTAIN',error=?,updated_at=? WHERE job_id=? AND status IN ('PENDING','SENDING')",
                         ("Fila recuperada de backup antigo. Confira no Gmail o que já foi enviado antes de confirmar ou liberar novas tentativas.", now, job["id"]))
            api._send_queue.refresh(conn, job["id"], finished=True)
            conn.execute("UPDATE send_jobs SET status='FAILED',message=?,current_label=NULL WHERE id=?",
                         ("Fila desativada após restauração. A restauração do banco não desfaz os e-mails enviados; confira os resultados no Gmail.", job["id"]))
    return sum(bool(record[field + "_sent_at"]) for record in records.values() for field in ("fixed", "complementary")), len(jobs)


def restore(api, source):
    """Restaura pela API SQLite sem remover WAL/SHM e sem repetir SMTP passado."""
    with job_lock(api.DATA_DIR, "database-maintenance") as acquired:
        if not acquired:
            raise ValueError("Outra operação está preparando a base. Aguarde e tente novamente.")
        with closing(api.db()) as conn:
            reason = api.delivery_mutation_reason(conn)
        if reason:
            raise ValueError(reason)
        return _restore_owned(api, Path(source))


def _restore_owned(api, source):
    staging = api.DATA_DIR / ("restore_" + uuid.uuid4().hex)
    staging.mkdir()
    previous, applied = None, []
    database_changed = False
    try:
        staged_db, staged_zip = staging / "snapshot.db", staging / "snapshot.zip"
        shutil.copy2(source, staged_db)
        if source.with_suffix(".zip").is_file():
            shutil.copy2(source.with_suffix(".zip"), staged_zip)
        assets, keys = _validate_source(api, staged_db, staged_zip, staging)
        records, logs, preserved_docs, missing = _capture_external_knowledge(api, staging)
        previous = create(api, "antes_restauracao", allow_missing_assets=True)

        def install_asset(staged, destination):
            destination.parent.mkdir(parents=True, exist_ok=True)
            original = None
            if destination.is_file():
                original = staging / "originals" / uuid.uuid4().hex
                original.parent.mkdir(exist_ok=True)
                shutil.copy2(destination, original)
            applied.append((destination, original))
            shutil.copy2(staged, destination)

        for staged, destination in assets:
            install_asset(staged, destination)
        with closing(sqlite3.connect(staged_db)) as src, closing(api.db()) as dst:
            src.backup(dst)
        database_changed = True
        api.init_db()
        preserved, disabled_jobs = _merge_external_knowledge(api, records, logs, preserved_docs, install_asset)
        if ".fernet_key" in keys:
            api.fernet = api.Fernet(keys[".fernet_key"].read_bytes().strip())
        if ".flask_secret" in keys:
            api.app.secret_key = keys[".flask_secret"].read_text(encoding="utf-8").strip()
        api.audit_event("PRESERVAR_ENVIO_APOS_RESTAURACAO", "sistema", description="Conhecimento dos envios SMTP preservado; filas antigas desativadas.",
                        metadata={"charges_preserved": preserved, "disabled_jobs": disabled_jobs, "missing_current_documents": missing,
                                  "backup_missing_files": keys.get("_missing_assets", [])})
        return previous, bool(assets)
    except Exception:
        if database_changed and previous and previous.is_file():
            with closing(sqlite3.connect(previous)) as src, closing(api.db()) as dst:
                src.backup(dst)
        for destination, original in reversed(applied):
            if original:
                shutil.copy2(original, destination)
            else:
                destination.unlink(missing_ok=True)
        raise
    finally:
        if staging.resolve().is_relative_to(api.DATA_DIR.resolve()):
            shutil.rmtree(staging)
