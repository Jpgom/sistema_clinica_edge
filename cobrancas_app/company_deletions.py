"""Remoção de cadastros em lote sem apagar efeitos financeiros ou SMTP."""
from contextlib import closing
import hashlib
import json

from itsdangerous import BadSignature, URLSafeTimedSerializer
from send_queue import job_lock


def _ids(values):
    result = set()
    for value in values:
        if isinstance(value, bool) or not str(value).isdigit() or int(value) <= 0:
            raise ValueError("Seleção de empresas inválida. Recarregue a lista.")
        result.add(int(value))
    if not result:
        raise ValueError("Selecione ao menos uma empresa.")
    return sorted(result)


def _rows_for_ids(conn, table, field, ids):
    result = []
    for start in range(0, len(ids), 400):
        batch = ids[start:start + 400]
        marks = ",".join("?" for _ in batch)
        result.extend(dict(row) for row in conn.execute(f"SELECT * FROM {table} WHERE {field} IN ({marks}) ORDER BY id", batch))
    return sorted(result, key=lambda row: row["id"])


def select_ids(conn, mode, values=(), unit_id=None):
    if unit_id is not None and not conn.execute("SELECT 1 FROM units WHERE id=?", (unit_id,)).fetchone():
        raise ValueError("A unidade escolhida não existe.")
    if mode == "all":
        query = "SELECT id FROM companies" + (" WHERE unit_id=?" if unit_id is not None else "") + " ORDER BY id"
        ids = [row[0] for row in conn.execute(query, (unit_id,) if unit_id is not None else ())]
        return _ids(ids)
    if mode != "selected":
        raise ValueError("Escolha as empresas selecionadas ou todas do escopo informado.")
    return _ids(values)


def build_plan(api, conn, values, unit_id=None):
    ids = _ids(values)
    companies = _rows_for_ids(conn, "companies", "id", ids)
    if len(companies) != len(ids):
        raise ValueError("A lista de empresas mudou. Refaça a seleção antes de apagar.")
    if unit_id is not None and any(row["unit_id"] != unit_id for row in companies):
        raise ValueError("A seleção inclui empresas fora da unidade escolhida.")
    units = {row["id"]: row["name"] for row in conn.execute("SELECT id,name FROM units")}
    company_units = {row["unit_id"] for row in companies}
    for uid in sorted(company_units & units.keys()):
        reason = api.delivery_mutation_reason(conn, unit_id=uid)
        if reason:
            raise ValueError(f"{units[uid]}: {reason}")
    if company_units - units.keys():
        reason = api.delivery_mutation_reason(conn)
        if reason:
            raise ValueError(f"Cadastro legado sem unidade válida: {reason}")
    charges = _rows_for_ids(conn, "competency_companies", "company_id", ids)
    logs = _rows_for_ids(conn, "email_logs", "company_id", ids)
    documents = _rows_for_ids(conn, "documents", "company_id", ids)
    exams = _rows_for_ids(conn, "exam_items", "company_id", ids)
    prices = _rows_for_ids(conn, "complementary_prices", "company_id", ids)
    competencies = _rows_for_ids(conn, "competencies", "id", sorted({row["competency_id"] for row in charges + documents + exams + logs}))
    closed = {row["id"] for row in competencies if row.get("status") == "FECHADA"}
    historical = {}
    for row in charges:
        if row["fixed_paid"] or row["complementary_paid"] or row["fixed_paid_at"] or row["complementary_paid_at"]:
            historical[row["company_id"]] = "Pagamento registrado"
        elif row["fixed_sent_at"] or row["complementary_sent_at"]:
            historical.setdefault(row["company_id"], "Cobrança já enviada")
        elif row["competency_id"] in closed:
            historical.setdefault(row["company_id"], "Participa de competência fechada")
    for row in logs:
        historical.setdefault(row["company_id"], "Histórico de tentativa de envio preservado")
    queue_rows = []
    target_set = set(ids)
    for row in conn.execute("SELECT * FROM send_job_groups ORDER BY id"):
        record = dict(row)
        try:
            members = set(json.loads(record.get("member_ids_json") or "[]")) | {record["company_id"]}
        except (ValueError, TypeError):
            raise ValueError("Uma fila antiga contém participantes inconsistentes. Confira as filas antes de apagar empresas.")
        involved = members & target_set
        if involved:
            queue_rows.append(record)
            for cid in involved:
                historical.setdefault(cid, "Registro de fila e recibo de envio preservados")
    memberships = [
        dict(row) for row in conn.execute("SELECT group_id,company_id FROM billing_group_members ORDER BY group_id,company_id")
    ]
    groups = [dict(row) for row in conn.execute("SELECT * FROM billing_groups ORDER BY id")]
    impacted = []
    for group in groups:
        members = {row["company_id"] for row in memberships if row["group_id"] == group["id"]}
        removed = members & target_set
        if not removed:
            continue
        remaining = members - target_set
        dissolve = group["primary_company_id"] in target_set or len(remaining) < 2
        impacted.append({"id": group["id"], "name": group["name"], "action": "dissolve" if dissolve else "update", "removed": sorted(removed), "remaining": sorted(remaining)})
    targets = []
    for company in companies:
        reason = historical.get(company["id"])
        targets.append({"id": company["id"], "name": company["name"], "cnpj": company["cnpj"], "unit_name": units.get(company["unit_id"], "Sem unidade válida"), "unit_id": company["unit_id"], "active": company["active"],
                        "action": "archive" if reason else "delete", "reason": reason or "Cadastro sem emissão, pagamento ou tentativa de envio"})
    evidence = {"companies": companies, "charges": charges, "logs": logs, "documents": documents, "exams": exams, "prices": prices,
                "competencies": competencies, "queue_rows": queue_rows, "impacted_groups": impacted,
                "groups": [group for group in groups if group["id"] in {item["id"] for item in impacted}]}
    fingerprint = hashlib.sha256(json.dumps(evidence, sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode()).hexdigest()
    return {"ids": ids, "targets": targets, "total": len(ids), "delete_count": sum(item["action"] == "delete" for item in targets),
            "archive_count": sum(item["action"] == "archive" for item in targets), "active_count": sum(bool(item["active"]) for item in targets),
            "already_inactive": sum(item["action"] == "archive" and not item["active"] for item in targets),
            "groups": impacted, "fingerprint": fingerprint, "unit_id": unit_id,
            "scope_label": units[unit_id] if unit_id is not None else "todas as unidades"}


def encode_plan(api, plan, mode="selected"):
    return URLSafeTimedSerializer(api.app.secret_key, salt="company-deletion-review").dumps({"ids": plan["ids"], "unit_id": plan["unit_id"], "fingerprint": plan["fingerprint"], "mode": mode})


def decode_plan(api, token):
    try:
        result = URLSafeTimedSerializer(api.app.secret_key, salt="company-deletion-review").loads(token, max_age=1800)
        result["ids"] = _ids(result["ids"])
        return result
    except (BadSignature, KeyError, TypeError, ValueError) as exc:
        raise ValueError("A revisão expirou ou está incompleta. Selecione as empresas e revise novamente.") from exc


def execute(api, values, unit_id=None, expected_fingerprint=None):
    """Revalida e aplica um lote completo; arquivos só saem após commit."""
    with job_lock(api.DATA_DIR, "database-maintenance", timeout=10) as acquired:
        if not acquired:
            raise ValueError("Outra operação está preparando a base. Aguarde e revise novamente.")
        with closing(api.db()) as conn:
            try:
                conn.execute("BEGIN IMMEDIATE")
                plan = build_plan(api, conn, values, unit_id)
                if expected_fingerprint and plan["fingerprint"] != expected_fingerprint:
                    raise ValueError("Os cadastros, cobranças ou grupos mudaram desde a revisão. Revise novamente antes de apagar.")
                backup = api.create_db_backup(f"antes_apagar_empresas_lote_{plan['total']}")
                stamp = api.now_iso()
                for group in plan["groups"]:
                    if group["action"] == "dissolve":
                        conn.execute("DELETE FROM billing_groups WHERE id=?", (group["id"],))
                    else:
                        conn.executemany("DELETE FROM billing_group_members WHERE group_id=? AND company_id=?", [(group["id"], cid) for cid in group["removed"]])
                        conn.execute("UPDATE billing_groups SET updated_at=? WHERE id=?", (stamp, group["id"]))
                deleted_documents = []
                competencies = set()
                for target in plan["targets"]:
                    cid = target["id"]
                    competencies.update(row[0] for row in conn.execute("SELECT competency_id FROM competency_companies WHERE company_id=?", (cid,)))
                    if target["action"] == "archive":
                        conn.execute("UPDATE companies SET active=0,updated_at=? WHERE id=?", (stamp, cid))
                    else:
                        deleted_documents.extend(conn.execute("SELECT stored_name FROM documents WHERE company_id=?", (cid,)).fetchall())
                        for table in ("documents", "exam_items", "competency_companies", "complementary_prices"):
                            conn.execute(f"DELETE FROM {table} WHERE company_id=?", (cid,))
                        conn.execute("DELETE FROM companies WHERE id=?", (cid,))
                cleanup = [row for row in deleted_documents if not conn.execute("SELECT 1 FROM documents WHERE stored_name=? LIMIT 1", (row["stored_name"],)).fetchone()]
                metadata = {"ids": plan["ids"], "deleted": plan["delete_count"], "archived": plan["archive_count"], "already_inactive": plan["already_inactive"],
                            "groups": plan["groups"], "backup": backup.name}
                conn.execute("""INSERT INTO audit_logs(action,entity_type,entity_id,description,metadata_json,created_at) VALUES(?,?,?,?,?,?)""",
                             ("APAGAR_EMPRESAS_LOTE", "cadastro", "lote", f"{plan['delete_count']} empresa(s) removida(s), {plan['archive_count']} arquivada(s) com histórico preservado.", json.dumps(metadata, ensure_ascii=False), stamp))
                conn.commit()
            except Exception:
                conn.rollback()
                raise
        warnings = []
        try:
            api.remove_document_files(cleanup)
        except Exception:
            warnings.append("Os cadastros foram removidos, mas alguns arquivos sem vínculo precisam de limpeza posterior.")
        for competency_id in sorted(competencies):
            try:
                api.recompute_competency_status(competency_id)
            except Exception:
                warnings.append("Uma competência precisa atualizar seu estado na próxima conferência.")
        plan["backup"] = backup
        plan["warnings"] = list(dict.fromkeys(warnings))
        return plan
