"""Fila local persistente, com aceite SMTP e recuperação conservadora."""
from __future__ import annotations

import hashlib
import json
import os
import threading
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

from mail_transport import DeliveryUncertain, validate_config
import complementary_report

EMAIL_TYPES = {"FIXED", "COMPLEMENTARY", "COMBINED", "FIXED_REMINDER", "COMPLEMENTARY_REMINDER"}
TERMINAL_JOB_STATES = {"COMPLETED", "COMPLETED_WITH_ERRORS", "FAILED"}
GROUP_LABELS = {
    "PENDING": "Aguardando", "SENDING": "Enviando", "SENT": "Enviado",
    "TEST_SENT": "Teste enviado", "SKIPPED": "Ignorado após revalidação",
    "ERROR": "Erro; pode repetir", "UNCERTAIN": "Resultado incerto; confira o Gmail",
    "PARTIAL": "Envio parcial; confira os destinatários recusados", "TEST_PARTIAL": "Teste parcial",
}


def migrate_send_queue(conn):
    """Migração aditiva: preserva o histórico existente e as filas antigas."""
    columns = {r[1] for r in conn.execute("PRAGMA table_info(send_jobs)")}
    for name, definition in {
        "force": "INTEGER NOT NULL DEFAULT 0", "test_mode": "INTEGER", "test_email": "TEXT",
    }.items():
        if name not in columns:
            conn.execute(f"ALTER TABLE send_jobs ADD COLUMN {name} {definition}")
    columns = {r[1] for r in conn.execute("PRAGMA table_info(send_job_groups)")}
    for name in ("member_ids_json", "delivery_result_json"):
        if name not in columns:
            conn.execute(f"ALTER TABLE send_job_groups ADD COLUMN {name} TEXT")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_send_jobs_active ON send_jobs(competency_id,status)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_send_groups_status ON send_job_groups(job_id,status)")


@contextmanager
def job_lock(data_dir, job_id, timeout=0):
    """O sistema operacional solta a trava inclusive se o processo for encerrado."""
    lock_dir = Path(data_dir) / "queue_locks"
    lock_dir.mkdir(parents=True, exist_ok=True)
    filename = hashlib.sha256(str(job_id).encode()).hexdigest() + ".lock"
    handle = (lock_dir / filename).open("a+b")
    acquired = False
    try:
        if os.name == "nt":
            import msvcrt
            if handle.seek(0, 2) == 0:
                handle.write(b"0")
                handle.flush()
            handle.seek(0)
            deadline = time.monotonic() + timeout
            while True:
                try:
                    msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                    acquired = True
                    break
                except OSError:
                    if time.monotonic() >= deadline:
                        break
                    time.sleep(0.02)
        else:
            import fcntl
            deadline = time.monotonic() + timeout
            while True:
                try:
                    fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    acquired = True
                    break
                except OSError:
                    if time.monotonic() >= deadline:
                        break
                    time.sleep(0.02)
        yield acquired
    finally:
        if acquired:
            if os.name == "nt":
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
        handle.close()


class SendQueue:
    def __init__(self, services):
        # globals() permanece atualizado, permitindo preservar as funções públicas do app.
        self.services = services
        self._threads = {}
        self._threads_lock = threading.Lock()

    def __getattr__(self, name):
        try:
            return self.services[name]
        except KeyError:
            raise AttributeError(name) from None

    @staticmethod
    def email_type(value):
        value = str(value or "").strip().upper()
        if value not in EMAIL_TYPES:
            raise ValueError("Tipo de cobrança inválido.")
        return value

    def member_ids(self, group):
        return json.loads(group["member_ids_json"]) if group["member_ids_json"] else [group["company_id"]]

    def package_label(self, group, default):
        result = json.loads(group["delivery_result_json"] or "{}")
        snapshot = result.get("pending_snapshot", {})
        if snapshot.get("delivery_kind") in {"recipient_batch", "combined"}:
            return f"{len(self.member_ids(group))} empresas → {snapshot.get('to_email') or result.get('to_email') or ''}"
        return default

    def uncertain_members(self, conn, competency_id, email_type):
        billing = email_type.split("_", 1)[0]
        rows = conn.execute(
            """SELECT g.company_id,g.member_ids_json,g.status,g.delivery_result_json FROM send_job_groups g JOIN send_jobs j ON j.id=g.job_id
               WHERE j.competency_id=? AND (?='COMBINED' OR j.email_type='COMBINED' OR j.email_type=? OR j.email_type=?) AND g.status IN ('UNCERTAIN','PARTIAL','TEST_PARTIAL')""",
            (competency_id, billing, billing, billing + "_REMINDER"),
        ).fetchall()
        return {cid for group in rows if group["status"] == "UNCERTAIN" or not json.loads(group["delivery_result_json"] or "{}").get("primary_accepted")
                for cid in self.member_ids(group)}

    def eligible(self, conn, competency_id, email_type):
        email_type = self.email_type(email_type)
        blocked = self.uncertain_members(conn, competency_id, email_type)
        ids = self._eligible_billing_company_ids(conn, competency_id, email_type)
        return [cid for cid in ids if not blocked.intersection(
            self.billing_delivery_ids(conn, competency_id, email_type, cid))]

    def reservation_conflict(self, conn, job, members):
        billing = job["email_type"].split("_", 1)[0]
        # Filas legadas podem ter sido criadas simultaneamente. O primeiro item
        # SENDING reserva os participantes, inclusive entre cobrança e lembrete.
        rows = conn.execute(
            """SELECT g.company_id,g.member_ids_json FROM send_job_groups g JOIN send_jobs j ON j.id=g.job_id
               WHERE j.competency_id=? AND j.id!=? AND (?='COMBINED' OR j.email_type='COMBINED' OR j.email_type=? OR j.email_type=?)
               AND g.status IN ('SENDING','UNCERTAIN')""",
            (job["competency_id"], job["id"], billing, billing, billing + "_REMINDER"),
        ).fetchall()
        return any(set(members).intersection(self.member_ids(other)) for other in rows)

    def payload_snapshot(self, conn, competency_id, email_type, company_id, members, payload=None):
        payload = payload or self.build_email(conn, competency_id, company_id, email_type, company_ids=members)
        if not payload:
            raise ValueError("Dados da cobrança não encontrados para criar a fila.")
        field = "fixed_amount" if email_type.startswith("FIXED") else "complementary_amount"
        values = {str(cid): conn.execute(
            f"SELECT {field} FROM competency_companies WHERE competency_id=? AND company_id=?",
            (competency_id, cid)).fetchone()[0] for cid in members}
        if email_type == 'COMBINED':
            values = {cid: round(sum(parts.values()), 2) for cid, parts in payload['financial_components'].items()}
        attachment_ids = payload.get("attachment_company_ids", [company_id])
        docs = [{key: row[key] for key in ('id','company_id','original_name','stored_name','sha256')}
                for row in self.payload_documents(conn, competency_id, email_type, company_id, payload)]
        content = json.dumps([payload["subject"], payload["html"], payload["text"]], ensure_ascii=False)
        snapshot = {
            "company_amounts": values, "to_email": payload["row"]["email"],
            "email_cc": payload["row"]["email_cc"] or "", "documents": docs,
            "content_sha256": hashlib.sha256(content.encode("utf-8")).hexdigest(),
            "delivery_kind": payload.get("delivery_kind", "explicit_group" if payload.get("grouped") else "individual"),
            "attachment_company_ids": attachment_ids,
        }
        if email_type == 'COMBINED':
            snapshot['financial_components'] = payload['financial_components']
            snapshot['document_refs'] = payload['document_refs']
        reports = [a['report_sha256'] for a in payload['attachments'] if a.get('generated') == 'complementary_report']
        if reports:
            snapshot['complementary_reports'] = reports
        return snapshot

    def payload_documents(self, conn, competency_id, email_type, company_id, payload):
        refs = payload.get('document_refs')
        if refs is None:
            refs = [(cid, email_type.split('_', 1)[0]) for cid in payload.get('attachment_company_ids', [company_id])]
        docs = []
        for cid, billing in refs:
            docs.extend(conn.execute('SELECT * FROM documents WHERE competency_id=? AND company_id=? AND billing_type=? ORDER BY id',
                                     (competency_id, cid, billing)).fetchall())
        return sorted(docs, key=lambda d: (d['company_id'], d['id']))

    def dispatch(self, job_id):
        with self._threads_lock:
            previous = self._threads.get(job_id)
            if previous and previous.is_alive():
                return
            thread = threading.Thread(target=self.run, args=(job_id,), daemon=True, name=f"billing-{job_id[:8]}")
            self._threads[job_id] = thread
            thread.start()

    def recover(self):
        if os.environ.get("ENVIO_COBRANCAS_DISABLE_JOB_RECOVERY") == "1":
            return 0
        conn = self.db()
        try:
            ids = [r["id"] for r in conn.execute(
                "SELECT id FROM send_jobs WHERE status IN ('QUEUED','RUNNING') ORDER BY created_at,id")]
        finally:
            conn.close()
        for job_id in ids:
            self.dispatch(job_id)
        return len(ids)

    def create(self, competency_id, email_type, ids, parent_job_id=None, force=False):
        with job_lock(self.DATA_DIR, "database-maintenance", timeout=10) as acquired:
            if not acquired:
                raise ValueError("A base está sendo restaurada ou preparada. Aguarde antes de iniciar o envio.")
            return self._create(competency_id, email_type, ids, parent_job_id, force)

    def _create(self, competency_id, email_type, ids, parent_job_id=None, force=False):
        email_type = self.email_type(email_type)
        cfg = self.smtp_config()
        validate_config(cfg, self.valid_email)
        unique_ids = []
        for value in ids:
            if isinstance(value, bool) or not str(value).isdigit() or int(value) <= 0:
                raise ValueError("Identificador de empresa inválido.")
            if int(value) not in unique_ids:
                unique_ids.append(int(value))
        if not unique_ids:
            raise ValueError("Nenhuma empresa está pronta para este envio. Confira também resultados incertos nas filas anteriores.")
        conn = self.db()
        try:
            conn.execute("BEGIN IMMEDIATE")
            comp = conn.execute("SELECT * FROM competencies WHERE id=?", (competency_id,)).fetchone()
            if parent_job_id and not conn.execute("SELECT id FROM send_jobs WHERE id=?", (parent_job_id,)).fetchone():
                raise ValueError("A fila anterior não está mais nesta base. Revise as cobranças e inicie uma nova fila.")
            if not comp:
                raise ValueError("Competência não encontrada.")
            if self.competency_is_closed(comp) and not email_type.endswith("_REMINDER"):
                raise ValueError("A competência está fechada. Reabra para fazer novos envios.")
            active = conn.execute(
                "SELECT id,email_type FROM send_jobs WHERE competency_id=? AND status IN ('QUEUED','RUNNING') ORDER BY created_at LIMIT 1",
                (competency_id,),
            ).fetchone()
            if active:
                if active["email_type"] != email_type:
                    raise ValueError("Aguarde a fila em andamento desta competência terminar antes de iniciar outro tipo de envio.")
                job_id, created = active["id"], False
            else:
                snapshots = []
                blocked = self.uncertain_members(conn, competency_id, email_type)
                representatives = set()
                for selected_cid in unique_ids:
                    cid = self.billing_delivery_representative(conn, competency_id, email_type, selected_cid, force=force)
                    if cid in representatives:
                        continue
                    representatives.add(cid)
                    reason = self.billing_send_reason(conn, competency_id, email_type, cid, force=force)
                    if reason:
                        raise ValueError(reason)
                    members = self.billing_delivery_ids(conn, competency_id, email_type, cid, force=force)
                    if not members:
                        raise ValueError("Nenhuma empresa elegível neste pacote de envio.")
                    if blocked.intersection(members):
                        raise ValueError("Existe envio com resultado incerto para esta cobrança. Confira o Gmail e resolva a fila anterior antes de reenviar.")
                    if any(set(members).intersection(previous) for _, previous, _ in snapshots):
                        raise ValueError("Uma empresa aparece em mais de um pacote de envio.")
                    snapshot = self.payload_snapshot(conn, competency_id, email_type, cid, members)
                    snapshots.append((cid, members, snapshot))
                job_id, created = uuid.uuid4().hex, True
                conn.execute(
                    """INSERT INTO send_jobs(id,competency_id,email_type,status,total_groups,created_at,message,parent_job_id,force,test_mode,test_email)
                       VALUES(?,?,?,'QUEUED',?,?,?,?,?,?,?)""",
                    (job_id, competency_id, email_type, len(snapshots), self.now_iso(), "Fila criada.", parent_job_id,
                     int(bool(force)), int(bool(cfg["test_mode"])), cfg.get("test_email")),
                )
                for cid, members, snapshot in snapshots:
                    conn.execute(
                        """INSERT INTO send_job_groups(job_id,company_id,status,created_at,updated_at,member_ids_json,delivery_result_json)
                           VALUES(?,?,'PENDING',?,?,?,?)""",
                        (job_id, cid, self.now_iso(), self.now_iso(), json.dumps(members), json.dumps({"pending_snapshot": snapshot}, ensure_ascii=False)),
                    )
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
        if created:
            self.audit_event("CRIAR_FILA_ENVIO", "send_job", job_id,
                             f"Fila {email_type} criada com {len(snapshots)} pacote(s) de envio.",
                             {"parent_job_id": parent_job_id, "force": bool(force), "test_mode": bool(cfg["test_mode"])},
                             competency_id=competency_id)
        self.dispatch(job_id)
        return job_id, created

    def retry(self, job_id):
        conn = self.db()
        try:
            job = conn.execute("SELECT * FROM send_jobs WHERE id=?", (job_id,)).fetchone()
            if not job:
                raise ValueError("Envio anterior não encontrado.")
            if job["status"] not in TERMINAL_JOB_STATES:
                raise ValueError("Aguarde a fila anterior terminar antes de repetir os erros.")
            failed_groups = conn.execute(
                "SELECT * FROM send_job_groups WHERE job_id=? AND status='ERROR' ORDER BY id", (job_id,)).fetchall()
            ids = [r["company_id"] for r in failed_groups]
            if not ids:
                raise ValueError("Esta fila não possui erros que possam ser repetidos. Resultados incertos ou parciais exigem conferência no Gmail.")
            cfg = self.smtp_config()
            if job["test_mode"] is not None and bool(job["test_mode"]) != bool(cfg["test_mode"]):
                raise ValueError("O modo de teste mudou. Inicie uma nova fila no modo atual após revisar as cobranças.")
            for group in failed_groups:
                members = self.billing_delivery_ids(conn, job["competency_id"], job["email_type"], group["company_id"], force=bool(job["force"]))
                if group["member_ids_json"] and set(members) != set(self.member_ids(group)):
                    raise ValueError("Os participantes ou pagamentos mudaram desde o erro. Revise a cobrança e inicie uma nova fila.")
                snapshot = json.loads(group["delivery_result_json"] or "{}").get("pending_snapshot")
                if snapshot and snapshot != self.payload_snapshot(conn, job["competency_id"], job["email_type"], group["company_id"], members):
                    raise ValueError("Os valores, destinatários ou documentos mudaram desde o erro. Revise a cobrança e inicie uma nova fila.")
        finally:
            conn.close()
        return self.create(job["competency_id"], job["email_type"], ids, parent_job_id=job_id, force=bool(job["force"]))

    def validated_attachments(self, conn, job, group, payload):
        docs = self.payload_documents(conn, job['competency_id'], job['email_type'], group['company_id'], payload)
        physical = [a for a in payload['attachments'] if not a.get('generated')]
        if not docs or len(docs) != len(physical):
            raise ValueError("Um ou mais anexos do pacote estão ausentes. Corrija os documentos antes de enviar.")
        filenames = {attachment.get("document_id"): attachment["filename"] for attachment in physical}
        attachments = []
        for doc in docs:
            path = Path(self.DATA_DIR) / doc["stored_name"]
            if not path.is_file():
                raise ValueError(f"Anexo não encontrado: {doc['original_name']}")
            data = path.read_bytes()
            if not data or hashlib.sha256(data).hexdigest() != doc["sha256"]:
                raise ValueError(f"Anexo vazio ou alterado: {doc['original_name']}. Adicione o documento novamente.")
            attachments.append({"path": str(path), "filename": filenames.get(doc["id"], doc["original_name"]), "data": data})
        for attachment in payload['attachments']:
            if not attachment.get('generated'):
                continue
            if attachment['generated'] != 'complementary_report' or complementary_report.digest(attachment['report']) != attachment['report_sha256']:
                raise ValueError('O demonstrativo dos exames mudou. Revise e crie uma nova fila.')
            attachments.append({'generated':'complementary_report','filename':attachment['filename'],
                'mime_type':complementary_report.MIME,'data':complementary_report.workbook_bytes(attachment['report'])})
        return attachments

    def counts(self, conn, job_id):
        values = {r["status"]: r["n"] for r in conn.execute(
            "SELECT status,COUNT(*) n FROM send_job_groups WHERE job_id=? GROUP BY status", (job_id,))}
        return {
            "sent": values.get("SENT", 0), "test_sent": values.get("TEST_SENT", 0),
            "skipped": values.get("SKIPPED", 0), "error": values.get("ERROR", 0),
            "partial": values.get("PARTIAL", 0) + values.get("TEST_PARTIAL", 0),
            "uncertain": values.get("UNCERTAIN", 0),
            "pending": values.get("PENDING", 0) + values.get("SENDING", 0),
        }

    def refresh(self, conn, job_id, finished=False):
        counts = self.counts(conn, job_id)
        processed = sum(counts[k] for k in counts if k != "pending")
        errors = counts["error"] + counts["partial"] + counts["uncertain"]
        conn.execute(
            "UPDATE send_jobs SET processed_groups=?,sent_groups=?,error_groups=?,heartbeat_at=? WHERE id=?",
            (processed, counts["sent"], errors, self.now_iso(), job_id),
        )
        if finished:
            status = "COMPLETED_WITH_ERRORS" if errors else "COMPLETED"
            message = (f"Concluído: {counts['sent']} enviados, {counts['test_sent']} testes, "
                       f"{counts['skipped']} ignorados, {counts['error']} erros, "
                       f"{counts['partial']} parciais e {counts['uncertain']} incertos.")
            conn.execute(
                "UPDATE send_jobs SET status=?,message=?,current_label=NULL,finished_at=?,heartbeat_at=? WHERE id=?",
                (status, message, self.now_iso(), self.now_iso(), job_id),
            )
        return counts

    def record_delivery(self, conn, job, group, payload, result, status, error=""):
        sent_at = self.now_iso()
        conn.execute("UPDATE send_job_groups SET status=?,error=?,updated_at=?,delivery_result_json=? WHERE id=?",
                     (status, error or None, sent_at, json.dumps(result, ensure_ascii=False), group["id"]))
        real_accepted = not result["test_mode"] and result["primary_accepted"]
        log_status = "TESTE_PARCIAL" if status == "TEST_PARTIAL" else "ENVIO_PARCIAL" if status == "PARTIAL" else "TESTE" if result["test_mode"] else "ENVIADO"
        for cid in self.member_ids(group):
            billing = job["email_type"].split("_", 1)[0]
            amount_field = "fixed_amount" if billing == "FIXED" else "complementary_amount"
            cc = conn.execute("SELECT * FROM competency_companies WHERE competency_id=? AND company_id=?",
                              (job["competency_id"], cid)).fetchone()
            if not cc:
                continue
            if job['email_type'] == 'COMBINED':
                components = result['financial_components'][str(cid)]
                for modality, amount in components.items():
                    if modality not in {'FIXED','COMPLEMENTARY'}:
                        raise ValueError('Modalidade inválida no registro do envio conjunto.')
                    if real_accepted:
                        field = 'fixed_sent_at' if modality == 'FIXED' else 'complementary_sent_at'
                        conn.execute(f'UPDATE competency_companies SET {field}=?,updated_at=? WHERE competency_id=? AND company_id=?',
                                     (sent_at,sent_at,job['competency_id'],cid))
                    conn.execute('''INSERT INTO email_logs(competency_id,company_id,email_type,to_email,subject,amount,status,error,sent_at,created_at)
                        VALUES(?,?,?,?,?,?,?,?,?,?)''', (job['competency_id'],cid,modality,result['to_email'],
                        result.get('subject',payload['subject']),amount,log_status,error or None,sent_at,sent_at))
                continue
            if real_accepted and job["email_type"] in {"FIXED", "COMPLEMENTARY"}:
                field = "fixed_sent_at" if billing == "FIXED" else "complementary_sent_at"
                conn.execute(f"UPDATE competency_companies SET {field}=?,updated_at=? WHERE competency_id=? AND company_id=?",
                             (sent_at, sent_at, job["competency_id"], cid))
            conn.execute(
                """INSERT INTO email_logs(competency_id,company_id,email_type,to_email,subject,amount,status,error,sent_at,created_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (job["competency_id"], cid, job["email_type"], result["to_email"], result.get("subject", payload["subject"]),
                 result.get("company_amounts", {}).get(str(cid), cc[amount_field]), log_status, error or None, sent_at, sent_at),
            )
        if real_accepted and job["email_type"] in {"FIXED", "COMPLEMENTARY", "COMBINED"}:
            conn.execute("UPDATE competencies SET status='ENVIADA',updated_at=? WHERE id=? AND COALESCE(status,'')!='FECHADA'",
                         (sent_at, job["competency_id"]))

    def record_problem(self, conn, job, group, status, error):
        stamp = self.now_iso()
        conn.execute("UPDATE send_job_groups SET status=?,error=?,updated_at=? WHERE id=?",
                     (status, error[:1000], stamp, group["id"]))
        if status != "SKIPPED":
            for cid in self.member_ids(group):
                conn.execute(
                    "INSERT INTO email_logs(competency_id,company_id,email_type,status,error,created_at) VALUES(?,?,?,?,?,?)",
                    (job["competency_id"], cid, job["email_type"], "INCERTO" if status == "UNCERTAIN" else "ERRO", error[:1000], stamp),
                )

    def run(self, job_id):
        try:
            with job_lock(self.DATA_DIR, job_id) as acquired:
                if acquired:
                    self._run_owned(job_id)
        except Exception as exc:
            self.app.logger.exception("Falha na execução da fila %s", job_id)
            try:
                conn = self.db()
                conn.execute("BEGIN IMMEDIATE")
                job = conn.execute("SELECT * FROM send_jobs WHERE id=?", (job_id,)).fetchone()
                if job and job["status"] in {"QUEUED", "RUNNING"}:
                    for group in conn.execute("SELECT * FROM send_job_groups WHERE job_id=? AND status IN ('PENDING','SENDING')", (job_id,)).fetchall():
                        self.record_problem(conn, job, group, "UNCERTAIN" if group["status"] == "SENDING" else "ERROR",
                                            "Fila interrompida. " + str(exc)[:500])
                    self.refresh(conn, job_id, finished=True)
                    conn.execute("UPDATE send_jobs SET status='FAILED',message=? WHERE id=?",
                                 (f"Fila interrompida: {str(exc)[:500]}. Confira os resultados antes de repetir os erros.", job_id))
                conn.commit()
                conn.close()
            except Exception:
                pass
        finally:
            with self._threads_lock:
                if self._threads.get(job_id) is threading.current_thread():
                    self._threads.pop(job_id, None)

    def _run_owned(self, job_id):
        conn = self.db()
        try:
            conn.execute("BEGIN IMMEDIATE")
            job = conn.execute("SELECT * FROM send_jobs WHERE id=?", (job_id,)).fetchone()
            if not job or job["status"] not in {"QUEUED", "RUNNING"}:
                conn.rollback()
                return
            self.email_type(job["email_type"])
            for group in conn.execute("SELECT * FROM send_job_groups WHERE job_id=? AND status='SENDING'", (job_id,)).fetchall():
                self.record_problem(conn, job, group, "UNCERTAIN",
                                    "Aplicativo interrompido durante o envio. Confira o Gmail e confirme o resultado antes de reenviar.")
            if job["test_mode"] is None:
                # A versão antiga não registrava o modo. Usar a configuração atual
                # poderia transformar uma fila de teste em cobrança real.
                for group in conn.execute("SELECT * FROM send_job_groups WHERE job_id=? AND status='PENDING'", (job_id,)).fetchall():
                    self.record_problem(conn, job, group, "ERROR", "Fila criada em versão anterior sem registro do modo de envio. Revise a configuração e as cobranças antes de iniciar nova tentativa.")
                self.refresh(conn, job_id, finished=True)
                conn.commit()
                return
            conn.execute("UPDATE send_jobs SET status='RUNNING',started_at=COALESCE(started_at,?),finished_at=NULL,heartbeat_at=?,message=? WHERE id=?",
                         (self.now_iso(), self.now_iso(), "Preparando envios e revalidando cobranças...", job_id))
            self.refresh(conn, job_id)
            conn.commit()
            job = conn.execute("SELECT * FROM send_jobs WHERE id=?", (job_id,)).fetchone()
            groups = conn.execute("SELECT * FROM send_job_groups WHERE job_id=? AND status='PENDING' ORDER BY id", (job_id,)).fetchall()
            for group in groups:
                payload = None
                transmission_started = False
                accepted_result = None
                try:
                    conn.execute("BEGIN IMMEDIATE")
                    force = bool(job["force"])
                    comp = conn.execute("SELECT * FROM competencies WHERE id=?", (job["competency_id"],)).fetchone()
                    reason = None
                    if not comp or (self.competency_is_closed(comp) and not job["email_type"].endswith("_REMINDER")):
                        reason = "Competência fechada ou excluída durante a fila."
                    if not reason:
                        reason = self.billing_send_reason(conn, job["competency_id"], job["email_type"], group["company_id"], force=force)
                    members = self.billing_delivery_ids(conn, job["competency_id"], job["email_type"], group["company_id"], force=force) if not reason else []
                    if not reason and group["member_ids_json"] and set(members) != set(self.member_ids(group)):
                        reason = "As empresas do pacote, pagamentos ou valores elegíveis mudaram durante a fila. Revise e inicie uma nova fila."
                    if not reason and self.reservation_conflict(conn, job, members):
                        reason = "Outra fila está enviando esta cobrança ou possui resultado incerto. Confira a fila anterior antes de reenviar."
                    if reason:
                        self.record_problem(conn, job, group, "SKIPPED", reason)
                        self.refresh(conn, job_id)
                        conn.commit()
                        continue
                    if not group["member_ids_json"]:
                        conn.execute("UPDATE send_job_groups SET member_ids_json=? WHERE id=?", (json.dumps(members), group["id"]))
                        group = conn.execute("SELECT * FROM send_job_groups WHERE id=?", (group["id"],)).fetchone()
                    payload = self.build_email(conn, job["competency_id"], group["company_id"], job["email_type"], company_ids=members)
                    if not payload:
                        raise ValueError("Empresa não encontrada na competência.")
                    original_snapshot = json.loads(group["delivery_result_json"] or "{}").get("pending_snapshot")
                    current_snapshot = self.payload_snapshot(conn, job["competency_id"], job["email_type"], group["company_id"], members, payload)
                    if original_snapshot and original_snapshot != current_snapshot:
                        self.record_problem(conn, job, group, "SKIPPED", "Os valores, destinatários ou documentos mudaram após a criação da fila. Revise a cobrança e inicie uma nova fila.")
                        self.refresh(conn, job_id)
                        conn.commit()
                        continue
                    attachments = self.validated_attachments(conn, job, group, payload)
                    cfg = self.smtp_config()
                    cfg["test_mode"] = bool(job["test_mode"])
                    cfg["test_email"] = job["test_email"]
                    validate_config(cfg, self.valid_email)
                    amount_field = "fixed_amount" if job["email_type"].startswith("FIXED") else "complementary_amount"
                    planned_result = {
                        "pending_snapshot": original_snapshot or current_snapshot,
                        "test_mode": cfg["test_mode"],
                        "to_email": cfg["test_email"] if cfg["test_mode"] else payload["row"]["email"],
                        "subject": payload["subject"],
                        "company_amounts": {str(cid): conn.execute(
                            f"SELECT {amount_field} FROM competency_companies WHERE competency_id=? AND company_id=?",
                            (job["competency_id"], cid)).fetchone()[0] for cid in members},
                    }
                    if job['email_type'] == 'COMBINED':
                        planned_result['financial_components'] = payload['financial_components']
                        planned_result['company_amounts'] = current_snapshot['company_amounts']
                    conn.execute("UPDATE send_job_groups SET status='SENDING',error=NULL,updated_at=?,delivery_result_json=? WHERE id=? AND status='PENDING'",
                                 (self.now_iso(), json.dumps(planned_result, ensure_ascii=False), group["id"]))
                    conn.execute("UPDATE send_jobs SET current_label=?,heartbeat_at=? WHERE id=?",
                                 (payload.get("delivery_label", payload["row"]["name"]), self.now_iso(), job_id))
                    conn.commit()
                    transmission_started = True
                    accepted_result = self.smtp_send(payload["row"]["email"], payload["row"]["email_cc"] or "",
                                                     payload["subject"], payload["html"], payload["text"], attachments, config=cfg)
                    accepted_result["company_amounts"] = planned_result["company_amounts"]
                    accepted_result["pending_snapshot"] = planned_result["pending_snapshot"]
                    if job['email_type'] == 'COMBINED':
                        accepted_result['financial_components'] = planned_result['financial_components']
                    if accepted_result["refused"]:
                        status = "TEST_PARTIAL" if accepted_result["test_mode"] else "PARTIAL"
                        error = "Destinatários recusados: " + "; ".join(f"{r}: {reason}" for r, reason in accepted_result["refused"].items())
                    else:
                        status = "TEST_SENT" if accepted_result["test_mode"] else "SENT"
                        error = ""
                    self.record_delivery(conn, job, group, payload, accepted_result, status, error)
                    self.refresh(conn, job_id)
                    conn.commit()
                except Exception as exc:
                    conn.rollback()
                    # Resultado aceito mas não persistido é incerto; jamais erro repetível.
                    status = "UNCERTAIN" if isinstance(exc, DeliveryUncertain) or accepted_result is not None else "ERROR"
                    if transmission_started and isinstance(exc, (KeyError, TypeError)):
                        status = "UNCERTAIN"
                    self.record_problem(conn, job, group, status, str(exc))
                    self.refresh(conn, job_id)
                    conn.commit()
            self.refresh(conn, job_id, finished=True)
            conn.commit()
            final = conn.execute("SELECT * FROM send_jobs WHERE id=?", (job_id,)).fetchone()
            self.audit_event("ENVIO_CONCLUIDO", "send_job", job_id, final["message"],
                             {"status": final["status"], "email_type": job["email_type"]}, competency_id=job["competency_id"])
        finally:
            conn.close()

    def resolve(self, job_id, group_id, action, original_test_mode=None):
        if action not in {"delivered", "not_delivered"}:
            raise ValueError("Confirme se o Gmail enviou a mensagem ou se ela não foi enviada.")
        conn = self.db()
        try:
            conn.execute("BEGIN IMMEDIATE")
            job = conn.execute("SELECT * FROM send_jobs WHERE id=?", (job_id,)).fetchone()
            group = conn.execute("SELECT * FROM send_job_groups WHERE job_id=? AND id=?", (job_id, group_id)).fetchone()
            if not job or not group:
                raise ValueError("Envio não encontrado.")
            if job["status"] not in TERMINAL_JOB_STATES or group["status"] not in {"UNCERTAIN", "PARTIAL", "TEST_PARTIAL"}:
                raise ValueError("Somente resultados incertos ou parciais de filas concluídas podem ser confirmados.")
            planned_result = json.loads(group["delivery_result_json"] or "{}")
            if action == "not_delivered":
                if group["status"] in {"PARTIAL", "TEST_PARTIAL"} and planned_result.get("primary_accepted"):
                    raise ValueError("O Gmail confirmou o aceite do destinatário principal. Confira as cópias recusadas; se precisar repetir, use o reenvio explícito da cobrança.")
                conn.execute("UPDATE send_job_groups SET status='ERROR',error=?,updated_at=? WHERE id=?",
                             ("Operador conferiu no Gmail e confirmou que não houve envio. Pode repetir.", self.now_iso(), group_id))
            else:
                if job["test_mode"] is None:
                    if original_test_mode not in {"0", "1"}:
                        raise ValueError("Esta fila antiga não registrou o modo. Informe se o envio confirmado foi REAL ou TESTE.")
                    conn.execute("UPDATE send_jobs SET test_mode=? WHERE id=?", (int(original_test_mode), job_id))
                    job = conn.execute("SELECT * FROM send_jobs WHERE id=?", (job_id,)).fetchone()
                payload = {"subject": planned_result.get("subject", "Envio confirmado após interrupção")}
                test_mode = bool(job["test_mode"])
                result = dict(planned_result, test_mode=test_mode, primary_accepted=True, subject=payload["subject"])
                if not result.get("to_email"):
                    owner = conn.execute("SELECT email FROM companies WHERE id=?", (group["company_id"],)).fetchone()
                    result["to_email"] = job["test_email"] if test_mode else owner["email"] if owner else ""
                self.record_delivery(conn, job, group, payload, result, "TEST_SENT" if test_mode else "SENT",
                                     "Envio confirmado manualmente após conferência no Gmail.")
            self.refresh(conn, job_id, finished=True)
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
        self.audit_event("CONFIRMAR_RESULTADO_ENVIO", "send_job", job_id,
                         "Resultado incerto conferido no Gmail.", {"group_id": group_id, "action": action}, competency_id=job["competency_id"])
