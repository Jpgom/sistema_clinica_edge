from __future__ import annotations

import os
import re
import io
import csv
import json
import html
import time
import uuid
import hashlib
import zipfile
import sqlite3
import secrets
import threading
import unicodedata
import smtplib
import shutil
import sys
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import billing_groups as grouped_billing
import delivery_batches
import combined_billing
import complementary_report
import payment_imports
import billing_policy
import control_apuration
import backups as data_backups
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from datetime import datetime, date
from collections import defaultdict

from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file, abort, jsonify
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from werkzeug.security import generate_password_hash, check_password_hash
from cryptography.fernet import Fernet
from mail_transport import send_email
from send_queue import SendQueue, migrate_send_queue, GROUP_LABELS, TERMINAL_JOB_STATES

APP_NAME = "EDGE - Envio de cobranças"
APP_VERSION = "V2.15 INTEGRADO"
BASE_DIR = Path(__file__).resolve().parent
BUNDLED_DATA_DIR = BASE_DIR / "data"
DATA_DIR = Path(os.environ.get("ENVIO_COBRANCAS_DATA_DIR") or os.environ.get("DATA_DIR") or BUNDLED_DATA_DIR)
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = Path(os.environ.get("ENVIO_COBRANCAS_DB_PATH") or os.environ.get("DB_PATH") or (DATA_DIR / "cobrancas.db"))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)

# Preserva a base que já existia na versão standalone ao integrar o módulo ao
# Sistema Interno. Em disco persistente novo, a cópia inicial leva banco, chaves,
# documentos e controles atuais. Bases persistentes já existentes nunca são sobrescritas.
_bundled_db = BUNDLED_DATA_DIR / "cobrancas.db"
if DB_PATH.resolve() != _bundled_db.resolve() and not DB_PATH.exists() and _bundled_db.exists():
    shutil.copy2(_bundled_db, DB_PATH)
for _secret_name in (".fernet_key", ".flask_secret"):
    _src = BUNDLED_DATA_DIR / _secret_name
    _dst = DATA_DIR / _secret_name
    if _src.exists() and not _dst.exists():
        shutil.copy2(_src, _dst)
for _folder_name in ("documents", "controls"):
    _src = BUNDLED_DATA_DIR / _folder_name
    _dst = DATA_DIR / _folder_name
    if _src.exists() and not _dst.exists():
        shutil.copytree(_src, _dst)

DOCS_DIR = DATA_DIR / "documents"
CONTROL_DIR = DATA_DIR / "controls"
BACKUP_DIR = DATA_DIR / "backups"


def stored_file_path(stored_name, control=False):
    """Resolve referências do banco em Windows ou Linux sem alterar a base existente."""
    name = str(stored_name or "").strip().replace("\\", "/")
    if not name:
        raise ValueError("Referência de arquivo vazia.")
    relative = Path(name)
    if relative.is_absolute() or ".." in relative.parts:
        raise ValueError("Referência de arquivo inválida.")
    if control and (not relative.parts or relative.parts[0] != "controls"):
        relative = Path("controls") / relative
    result = (DATA_DIR / relative).resolve()
    root = DATA_DIR.resolve()
    if result != root and root not in result.parents:
        raise ValueError("Referência de arquivo fora da pasta de dados.")
    return result
for p in (DOCS_DIR, CONTROL_DIR, BACKUP_DIR):
    p.mkdir(parents=True, exist_ok=True)

MONTHS = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO",
    7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}
# Tipos de ASO que não são cobrados como complementares.
OCCUPATIONAL_TYPES = {
    "ADMISSIONAL", "DEMISSIONAL", "PERIODICO", "PERIÓDICO", "RETORNO AO TRABALHO",
    "MUDANCA DE RISCOS OCUPACIONAIS", "MUDANÇA DE RISCOS OCUPACIONAIS",
}
# Catálogo padrão da EDGE. Somente tipos cadastrados e ativos entram na apuração.
# Hemograma e reticulócitos foram removidos do módulo de cobranças.
DEFAULT_COMPLEMENTARY_EXAMS = [
    "AUDIOMETRIA",
    "ESPIROMETRIA",
    "ACUIDADE VISUAL",
    "ANAMNESE PSICOSSOCIAL OCUPACIONAL",
    "LAUDO PCD",
]
FORBIDDEN_COMPLEMENTARY_EXAMS = {"HEMOGRAMA", "RETICULOCITOS"}
CANCEL_TERMS = (
    "NAO REALIZADO", "NÃO REALIZADO", "CANCELADO", "CANCELADA", "ESTORNO",
    "NAO ENVIAR", "NÃO ENVIAR", "NAO COBRAR", "NÃO COBRAR",
)
ALLOWED_DOC_EXT = {".pdf", ".png", ".jpg", ".jpeg", ".doc", ".docx", ".xls", ".xlsx"}

SECRET_FILE = DATA_DIR / ".flask_secret"
FERNET_FILE = DATA_DIR / ".fernet_key"

def _get_or_create_text_secret(path: Path, generator):
    if path.exists():
        return path.read_text(encoding="utf-8").strip()
    value = generator()
    path.write_text(value, encoding="utf-8")
    return value

FLASK_SECRET = os.environ.get("SECRET_KEY") or _get_or_create_text_secret(SECRET_FILE, lambda: secrets.token_hex(32))
FERNET_KEY = os.environ.get("FERNET_KEY") or _get_or_create_text_secret(FERNET_FILE, lambda: Fernet.generate_key().decode())
fernet = Fernet(FERNET_KEY.encode())

app = Flask(__name__)
app.secret_key = FLASK_SECRET
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("ENVIO_COBRANCAS_MAX_UPLOAD_MB", "120")) * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"


def now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def digits(value):
    return re.sub(r"\D", "", str(value or ""))


def valid_company_document(value):
    return len(digits(value)) in {11, 14}


def format_document(value):
    d = digits(value)
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    return d or "—"


def format_cnpj(value):
    # Compatibilidade interna: o campo histórico 'cnpj' agora armazena CNPJ OU CPF.
    return format_document(value)


def format_cpf(value):
    return format_document(value)


def normalize_text(value):
    s = str(value or "").strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s)


def norm_header(value):
    return re.sub(r"[^A-Z0-9]", "", normalize_text(value))


def canonical_control_exam(value, active_catalog=None):
    """Normaliza o exame do controle usando o catálogo ativo do sistema.

    O V2.8 limitava a leitura a cinco nomes fixos. A partir do V2.9, qualquer exame
    complementar cadastrado e ativo no sistema pode ser reconhecido. Os aliases
    históricos continuam aceitos para compatibilidade com os controles da EDGE.
    """
    key = normalize_text(value)
    aliases = {
        "AUIDIOMETRIA": "AUDIOMETRIA",
        "ANAMNESE PSICOSSOCIAL": "ANAMNESE PSICOSSOCIAL OCUPACIONAL",
    }
    key = normalize_text(aliases.get(key, key))
    if key.startswith("ANAMNESE PSICOSSOCIAL"):
        key = "ANAMNESE PSICOSSOCIAL OCUPACIONAL"

    if active_catalog is not None:
        # active_catalog: {exam_key_normalizado: nome_exibicao}
        return active_catalog.get(key) or active_catalog.get(norm_header(key))

    # Compatibilidade para chamadas antigas/testes sem catálogo explícito.
    fallback = {normalize_text(name): name for name in DEFAULT_COMPLEMENTARY_EXAMS}
    return fallback.get(key)


def money(value):
    try:
        v = float(value or 0)
    except Exception:
        v = 0.0
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse_money(value):
    try:
        return parse_money_strict(value) or 0.0
    except (ValueError, TypeError):
        return 0.0


def parse_money_strict(value):
    """Valores em centavos, com arredondamento comercial e rejeição de não finitos."""
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, bool):
        raise ValueError("Valor monetário inválido: use um número, não SIM/NÃO.")
    s = str(value).strip().replace("R$", "").replace("\u00a0", "").replace(" ", "")
    if not isinstance(value, (int, float, Decimal)):
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
    try:
        amount = Decimal(s)
        if not amount.is_finite() or amount < 0 or amount > Decimal("999999999999.99"):
            raise ValueError(f"Valor monetário inválido ou fora do intervalo: {value}")
        return float(amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    except (InvalidOperation, OverflowError) as exc:
        raise ValueError(f"Valor monetário inválido: {value}") from exc


def valid_email(value):
    return bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", (value or "").strip()))


def month_label(month, year):
    return f"{MONTHS.get(int(month), str(month))}/{year}"


def parse_excel_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        s = value.strip()
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
            try:
                return datetime.strptime(s[:10], fmt).date()
            except Exception:
                pass
    return None


def split_emails(value):
    parts = re.split(r"[;,/\n\r]+", str(value or ""))
    out = []
    for x in parts:
        e = x.strip().lower()
        if valid_email(e) and e not in out:
            out.append(e)
    return out


def db():
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA busy_timeout=10000")
    return conn


def create_db_backup(label="backup"):
    return data_backups.create(sys.modules[__name__],label)


def confirm_delete_request():
    return normalize_text(request.form.get("confirm_text")) == "EXCLUIR"


def audit_event(action, entity_type, entity_id=None, description="", metadata=None, company_id=None, competency_id=None):
    """Registra ações operacionais relevantes sem interromper o fluxo principal."""
    try:
        conn = db()
        conn.execute(
            """INSERT INTO audit_logs(action,entity_type,entity_id,company_id,competency_id,description,metadata_json,created_at)
               VALUES(?,?,?,?,?,?,?,?)""",
            (str(action), str(entity_type), str(entity_id or ""), company_id, competency_id, str(description or ""), json.dumps(metadata or {}, ensure_ascii=False), now_iso()),
        )
        conn.commit(); conn.close()
    except Exception:
        try: conn.close()
        except Exception: pass


def competency_is_closed(comp):
    try:
        return str(comp["status"] or "").upper() == "FECHADA"
    except Exception:
        return False


def ensure_competency_editable(competency_id):
    comp = get_competency(competency_id)
    if not comp:
        abort(404)
    if competency_is_closed(comp):
        flash("Esta competência está FECHADA. Reabra a competência para alterar apuração, documentos ou envios.", "error")
        return None
    return comp


def exam_fingerprint(company_id, employee, exam_key, exam_date, job_title=""):
    raw = "|".join([str(company_id), normalize_text(employee), normalize_text(exam_key), str(exam_date or "")])
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def recompute_competency_status(competency_id):
    """Mantém PREPARACAO → PRONTA → ENVIADA sem sobrescrever FECHADA."""
    conn = db()
    comp = conn.execute("SELECT * FROM competencies WHERE id=?", (competency_id,)).fetchone()
    if not comp:
        conn.close(); return
    if str(comp["status"] or "").upper() == "FECHADA":
        conn.close(); return
    row = conn.execute(
        """SELECT SUM(CASE WHEN fixed_sent_at IS NOT NULL THEN 1 ELSE 0 END) fs,
                  SUM(CASE WHEN complementary_sent_at IS NOT NULL THEN 1 ELSE 0 END) cs
           FROM competency_companies WHERE competency_id=?""", (competency_id,)
    ).fetchone()
    has_sent = int(row["fs"] or 0) + int(row["cs"] or 0) > 0
    control = bool(comp["control_stored_name"] or comp["processed_at"])
    conn.close()
    if has_sent:
        status = "EM_ANDAMENTO"
    elif control:
        try:
            snap = competency_operational_snapshot(competency_id)
            status = "PRONTA" if snap and snap["counts"]["blocking"] == 0 and snap["counts"]["ready"] > 0 else "PREPARACAO"
        except Exception:
            status = "PREPARACAO"
    else:
        status = "PREPARACAO"
    conn=db(); conn.execute("UPDATE competencies SET status=?,updated_at=? WHERE id=?", (status, now_iso(), competency_id)); conn.commit(); conn.close()



def remove_document_files(rows):
    for r in rows:
        stored = r["stored_name"] if isinstance(r, sqlite3.Row) else r
        if not stored:
            continue
        try:
            path = stored_file_path(stored)
            if DATA_DIR.resolve() in path.parents:
                path.unlink(missing_ok=True)
        except Exception:
            pass


def remove_control_file(stored_name):
    if not stored_name:
        return
    try:
        path = stored_file_path(stored_name, control=True)
        if CONTROL_DIR.resolve() in path.parents:
            path.unlink(missing_ok=True)
    except Exception:
        pass


def init_db():
    conn = db()
    conn.execute("PRAGMA journal_mode=WAL")
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS units (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unit_id INTEGER NOT NULL,
            cnpj TEXT NOT NULL,
            owner_cpf TEXT,
            name TEXT NOT NULL,
            email TEXT,
            email_cc TEXT,
            fixed_value REAL NOT NULL DEFAULT 0,
            bill_complementaries INTEGER NOT NULL DEFAULT 1,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(unit_id, cnpj),
            FOREIGN KEY(unit_id) REFERENCES units(id)
        );
        CREATE TABLE IF NOT EXISTS complementary_prices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            exam_name TEXT NOT NULL,
            exam_key TEXT NOT NULL,
            unit_price REAL NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(company_id, exam_key),
            FOREIGN KEY(company_id) REFERENCES companies(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS complementary_exam_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            exam_key TEXT NOT NULL UNIQUE,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS price_tables (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unit_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(unit_id, name),
            FOREIGN KEY(unit_id) REFERENCES units(id)
        );
        CREATE TABLE IF NOT EXISTS price_table_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            table_id INTEGER NOT NULL,
            exam_name TEXT NOT NULL,
            exam_key TEXT NOT NULL,
            unit_price REAL NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(table_id, exam_key),
            FOREIGN KEY(table_id) REFERENCES price_tables(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS competencies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unit_id INTEGER NOT NULL,
            month INTEGER NOT NULL,
            year INTEGER NOT NULL,
            control_filename TEXT,
            control_hash TEXT,
            control_stored_name TEXT,
            processed_at TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(unit_id, month, year),
            FOREIGN KEY(unit_id) REFERENCES units(id)
        );
        CREATE TABLE IF NOT EXISTS competency_companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            competency_id INTEGER NOT NULL,
            company_id INTEGER NOT NULL,
            fixed_amount REAL NOT NULL DEFAULT 0,
            fixed_amount_manual INTEGER NOT NULL DEFAULT 0,
            complementary_amount REAL NOT NULL DEFAULT 0,
            fixed_sent_at TEXT,
            complementary_sent_at TEXT,
            fixed_paid INTEGER NOT NULL DEFAULT 0,
            complementary_paid INTEGER NOT NULL DEFAULT 0,
            fixed_paid_at TEXT,
            complementary_paid_at TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(competency_id, company_id),
            FOREIGN KEY(competency_id) REFERENCES competencies(id) ON DELETE CASCADE,
            FOREIGN KEY(company_id) REFERENCES companies(id)
        );
        CREATE TABLE IF NOT EXISTS exam_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            competency_id INTEGER NOT NULL,
            company_id INTEGER NOT NULL,
            employee TEXT NOT NULL,
            exam_name TEXT NOT NULL,
            exam_key TEXT NOT NULL,
            exam_date TEXT,
            job_title TEXT,
            receipt TEXT,
            source_value REAL,
            unit_price REAL,
            total REAL,
            source_row INTEGER,
            status TEXT NOT NULL DEFAULT 'OK',
            note TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(competency_id) REFERENCES competencies(id) ON DELETE CASCADE,
            FOREIGN KEY(company_id) REFERENCES companies(id)
        );
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            competency_id INTEGER NOT NULL,
            company_id INTEGER NOT NULL,
            billing_type TEXT NOT NULL,
            original_name TEXT NOT NULL,
            stored_name TEXT NOT NULL,
            sha256 TEXT NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(competency_id, company_id, billing_type, sha256),
            FOREIGN KEY(competency_id) REFERENCES competencies(id) ON DELETE CASCADE,
            FOREIGN KEY(company_id) REFERENCES companies(id)
        );
        CREATE TABLE IF NOT EXISTS email_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            competency_id INTEGER NOT NULL,
            company_id INTEGER NOT NULL,
            email_type TEXT NOT NULL,
            to_email TEXT,
            subject TEXT,
            amount REAL,
            status TEXT NOT NULL,
            error TEXT,
            sent_at TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(competency_id) REFERENCES competencies(id) ON DELETE CASCADE,
            FOREIGN KEY(company_id) REFERENCES companies(id)
        );
        CREATE TABLE IF NOT EXISTS send_jobs (
            id TEXT PRIMARY KEY,
            competency_id INTEGER NOT NULL,
            email_type TEXT NOT NULL,
            status TEXT NOT NULL,
            total_groups INTEGER NOT NULL DEFAULT 0,
            processed_groups INTEGER NOT NULL DEFAULT 0,
            sent_groups INTEGER NOT NULL DEFAULT 0,
            error_groups INTEGER NOT NULL DEFAULT 0,
            current_label TEXT,
            message TEXT,
            created_at TEXT NOT NULL,
            started_at TEXT,
            finished_at TEXT,
            heartbeat_at TEXT,
            FOREIGN KEY(competency_id) REFERENCES competencies(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS send_job_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id TEXT NOT NULL,
            company_id INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'PENDING',
            error TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(job_id, company_id),
            FOREIGN KEY(job_id) REFERENCES send_jobs(id) ON DELETE CASCADE,
            FOREIGN KEY(company_id) REFERENCES companies(id)
        );
        """
    )
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            action TEXT NOT NULL,
            entity_type TEXT NOT NULL,
            entity_id TEXT,
            company_id INTEGER,
            competency_id INTEGER,
            description TEXT,
            metadata_json TEXT,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS apuration_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            competency_id INTEGER NOT NULL,
            original_name TEXT,
            file_hash TEXT,
            stats_json TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(competency_id) REFERENCES competencies(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS apuration_audit (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER NOT NULL,
            competency_id INTEGER NOT NULL,
            source_row INTEGER,
            decision TEXT NOT NULL,
            reason TEXT,
            document TEXT,
            company_id INTEGER,
            company_name TEXT,
            employee TEXT,
            source_exam TEXT,
            canonical_exam TEXT,
            exam_date TEXT,
            unit_price REAL,
            amount REAL,
            fingerprint TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(run_id) REFERENCES apuration_runs(id) ON DELETE CASCADE,
            FOREIGN KEY(competency_id) REFERENCES competencies(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS document_import_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            competency_id INTEGER NOT NULL,
            billing_type TEXT NOT NULL,
            zip_name TEXT,
            total_files INTEGER NOT NULL DEFAULT 0,
            added INTEGER NOT NULL DEFAULT 0,
            duplicates INTEGER NOT NULL DEFAULT 0,
            unmatched INTEGER NOT NULL DEFAULT 0,
            skipped INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            FOREIGN KEY(competency_id) REFERENCES competencies(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS document_import_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            company_id INTEGER,
            company_name TEXT,
            result TEXT NOT NULL,
            reason TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(run_id) REFERENCES document_import_runs(id) ON DELETE CASCADE
        );
        """
    )
    billing_policy.migrate(conn)
    company_cols = {r[1] for r in conn.execute("PRAGMA table_info(companies)").fetchall()}
    if "bill_complementaries" not in company_cols:
        # Migração segura: versões anteriores cobravam complementares de todas as empresas.
        conn.execute("ALTER TABLE companies ADD COLUMN bill_complementaries INTEGER NOT NULL DEFAULT 1")
    if "owner_cpf" not in company_cols:
        # CPF opcional do responsável/dono, usado apenas como identificador alternativo de documentos.
        conn.execute("ALTER TABLE companies ADD COLUMN owner_cpf TEXT")
    if "price_table_id" not in company_cols:
        # Tabela de preços reutilizável da V2. Mantém compatibilidade com os preços individuais existentes.
        conn.execute("ALTER TABLE companies ADD COLUMN price_table_id INTEGER")

    competency_company_cols = {r[1] for r in conn.execute("PRAGMA table_info(competency_companies)").fetchall()}
    if "fixed_amount_manual" not in competency_company_cols:
        # Permite alterar a mensalidade apenas na competência atual sem mudar o cadastro-base.
        conn.execute("ALTER TABLE competency_companies ADD COLUMN fixed_amount_manual INTEGER NOT NULL DEFAULT 0")

    competency_cols = {r[1] for r in conn.execute("PRAGMA table_info(competencies)").fetchall()}
    if "status" not in competency_cols:
        conn.execute("ALTER TABLE competencies ADD COLUMN status TEXT NOT NULL DEFAULT 'PREPARACAO'")
    if "closed_at" not in competency_cols:
        conn.execute("ALTER TABLE competencies ADD COLUMN closed_at TEXT")
    if "reopened_at" not in competency_cols:
        conn.execute("ALTER TABLE competencies ADD COLUMN reopened_at TEXT")
    exam_cols = {r[1] for r in conn.execute("PRAGMA table_info(exam_items)").fetchall()}
    if "fingerprint" not in exam_cols:
        conn.execute("ALTER TABLE exam_items ADD COLUMN fingerprint TEXT")
    if "duplicate_of_item_id" not in exam_cols:
        conn.execute("ALTER TABLE exam_items ADD COLUMN duplicate_of_item_id INTEGER")
    if "duplicate_override" not in exam_cols:
        conn.execute("ALTER TABLE exam_items ADD COLUMN duplicate_override INTEGER NOT NULL DEFAULT 0")
    send_job_cols = {r[1] for r in conn.execute("PRAGMA table_info(send_jobs)").fetchall()}
    if "parent_job_id" not in send_job_cols:
        conn.execute("ALTER TABLE send_jobs ADD COLUMN parent_job_id TEXT")

    # Índices operacionais da V2.5: aceleram apuração, auditoria, documentos e histórico.
    conn.executescript(
        """
        CREATE INDEX IF NOT EXISTS idx_exam_items_fingerprint ON exam_items(fingerprint);
        CREATE INDEX IF NOT EXISTS idx_exam_items_competency_status ON exam_items(competency_id, status);
        CREATE INDEX IF NOT EXISTS idx_apuration_audit_run ON apuration_audit(run_id);
        CREATE INDEX IF NOT EXISTS idx_apuration_audit_competency ON apuration_audit(competency_id);
        CREATE INDEX IF NOT EXISTS idx_doc_import_items_run ON document_import_items(run_id);
        CREATE INDEX IF NOT EXISTS idx_documents_comp_company_type ON documents(competency_id, company_id, billing_type);
        CREATE INDEX IF NOT EXISTS idx_audit_company ON audit_logs(company_id);
        CREATE INDEX IF NOT EXISTS idx_audit_competency ON audit_logs(competency_id);
        CREATE INDEX IF NOT EXISTS idx_email_logs_comp_company ON email_logs(competency_id, company_id);
        """
    )

    # Migração única da V2.1: instala o catálogo padrão em bases novas, acrescenta
    # LAUDO PCD em bases antigas e remove Hemograma/Reticulócitos dos preços futuros.
    # Depois da migração, o usuário continua livre para excluir qualquer tipo de exame.
    catalog_migrated = conn.execute("SELECT value FROM settings WHERE key='migration_v21_catalog'").fetchone()
    if not catalog_migrated:
        existing_count = conn.execute("SELECT COUNT(*) n FROM complementary_exam_types").fetchone()["n"]
        install_names = list(DEFAULT_COMPLEMENTARY_EXAMS) if not existing_count else ["LAUDO PCD"]
        for exam_name in install_names:
            key = normalize_text(exam_name)
            row = conn.execute("SELECT id FROM complementary_exam_types WHERE exam_key=?", (key,)).fetchone()
            if row:
                conn.execute("UPDATE complementary_exam_types SET name=?,active=1,updated_at=? WHERE id=?", (exam_name, now_iso(), row["id"]))
            else:
                conn.execute("INSERT INTO complementary_exam_types(name,exam_key,active,created_at,updated_at) VALUES(?,?,1,?,?)", (exam_name, key, now_iso(), now_iso()))
        for forbidden in FORBIDDEN_COMPLEMENTARY_EXAMS:
            key = normalize_text(forbidden)
            conn.execute("UPDATE complementary_exam_types SET active=0,updated_at=? WHERE exam_key=?", (now_iso(), key))
            conn.execute("DELETE FROM complementary_prices WHERE exam_key=?", (key,))
            conn.execute("DELETE FROM price_table_items WHERE exam_key=?", (key,))
        conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('migration_v21_catalog',?)", (now_iso(),))

    for name in ("BELÉM", "MACAPÁ"):
        conn.execute(
            "INSERT OR IGNORE INTO units(name,active,created_at,updated_at) VALUES(?,1,?,?)",
            (name, now_iso(), now_iso()),
        )
    grouped_billing.migrate_billing_groups(conn)
    migrate_send_queue(conn)
    payment_imports.migrate_payment_imports(conn)
    conn.commit()
    conn.close()


init_db()


def generate_csrf_token():
    token = session.get("_csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["_csrf_token"] = token
    return token


def validate_csrf():
    if request.method != "POST":
        return
    expected = session.get("_csrf_token")
    received = request.form.get("_csrf_token") or request.headers.get("X-CSRF-Token")
    if not expected or not received or not secrets.compare_digest(str(expected), str(received)):
        abort(403)


def setting_get(key, default=""):
    conn = db()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row["value"] if row else default


def setting_set(key, value):
    conn = db()
    conn.execute(
        "INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, str(value) if value is not None else ""),
    )
    conn.commit()
    conn.close()


def encrypt_secret(value):
    if not value:
        return ""
    return fernet.encrypt(value.encode()).decode()


def decrypt_secret(value):
    if not value:
        return ""
    try:
        return fernet.decrypt(value.encode()).decode()
    except Exception:
        return ""


def local_auth_enabled():
    return str(os.environ.get("EDGE_LOCAL_AUTH", "1")).strip().lower() not in {"0", "false", "no", "off"}


def parent_auth_enabled():
    return str(os.environ.get("ENVIO_COBRANCAS_PARENT_AUTH", "0")).strip().lower() in {"1", "true", "yes", "on"}


def _parent_user_status():
    """Retorna (autenticado, cargo) consultando a base principal quando disponível."""
    user_id = session.get("user_id")
    if not user_id:
        return False, "anonimo"
    lookup = app.config.get("EDGE_PARENT_USER_LOOKUP")
    if callable(lookup):
        try:
            user = lookup(user_id)
            if not user:
                return False, "anonimo"
            try:
                active = int(user["ativo"]) == 1
            except Exception:
                active = False
            try:
                cargo = str(user["cargo"] or "").strip().lower()
            except Exception:
                cargo = ""
            return active, (cargo or "anonimo")
        except Exception:
            # Falha fechada: uma área administrativa não deve ser liberada quando
            # não for possível confirmar o usuário na base principal.
            return False, "anonimo"
    cargo = str(session.get("cargo") or "").strip().lower()
    return True, (cargo or "anonimo")


def role():
    if parent_auth_enabled():
        _active, cargo = _parent_user_status()
        return cargo
    return "administrador"


@app.before_request
def protect_module():
    endpoint = request.endpoint or ""
    if endpoint.startswith("static") or endpoint == "healthz":
        return None
    if parent_auth_enabled():
        # DispatcherMiddleware preserva a sessão do Sistema EDGE. A checagem aqui
        # é obrigatória porque as rotas montadas não passam pelo before_request
        # da aplicação principal.
        if not session.get("user_id"):
            next_path = (request.script_root or "") + request.path
            return redirect("/login?next=" + next_path)
        active, cargo = _parent_user_status()
        if not active:
            session.clear()
            next_path = (request.script_root or "") + request.path
            return redirect("/login?next=" + next_path)
        if cargo != "admin":
            flash("Apenas administradores podem acessar Cobranças.", "error")
            return redirect("/")
        validate_csrf()
        return None
    if local_auth_enabled():
        configured = bool(setting_get("admin_password_hash"))
        if not configured and endpoint != "setup":
            return redirect(url_for("setup"))
        if configured and endpoint not in {"login", "setup"} and not session.get("authenticated"):
            return redirect(url_for("login"))
    validate_csrf()
    return None


def delivery_mutation_reason(conn, unit_id=None, competency_id=None, include_uncertain=True):
    conditions=[]; params=[]
    if unit_id is not None: conditions.append("cp.unit_id=?"); params.append(unit_id)
    if competency_id is not None: conditions.append("j.competency_id=?"); params.append(competency_id)
    where=(" AND "+" AND ".join(conditions)) if conditions else ""
    if conn.execute("SELECT 1 FROM send_jobs j JOIN competencies cp ON cp.id=j.competency_id WHERE j.status IN ('QUEUED','RUNNING')"+where+" LIMIT 1",params).fetchone():
        return "Aguarde a fila de envio terminar antes de alterar estes dados."
    if include_uncertain:
        rows=conn.execute("SELECT g.status,g.delivery_result_json FROM send_job_groups g JOIN send_jobs j ON j.id=g.job_id JOIN competencies cp ON cp.id=j.competency_id WHERE g.status IN ('SENDING','UNCERTAIN','PARTIAL','TEST_PARTIAL')"+where,params).fetchall()
        for row in rows:
            receipt=json.loads(row["delivery_result_json"] or "{}")
            if row["status"] in {"SENDING","UNCERTAIN"} or not receipt.get("primary_accepted"):
                return "Há um envio sem confirmação. Confira a fila e confirme o resultado antes de alterar ou excluir estes dados."
    return None


def payment_is_sending(conn, competency_id, company_id):
    rows=conn.execute("SELECT g.company_id,g.member_ids_json FROM send_job_groups g JOIN send_jobs j ON j.id=g.job_id WHERE j.competency_id=? AND g.status='SENDING'",(competency_id,)).fetchall()
    return any(company_id in (json.loads(r["member_ids_json"]) if r["member_ids_json"] else [r["company_id"]]) for r in rows)


@app.before_request
def preserve_operational_consistency():
    if request.method != "POST": return None
    endpoint=request.endpoint or ""
    company_endpoints={"company_new","company_edit","company_delete","company_price_delete"}
    price_endpoints={"price_table_new","price_table_edit","price_table_delete"}
    global_endpoints={"exam_types","exam_type_toggle","exam_type_delete","settings","settings_backup_restore"}
    competency_endpoints={"competency_close","competency_reopen","competency_delete"}
    if endpoint not in company_endpoints|price_endpoints|global_endpoints|competency_endpoints|{"unit_delete"}:return None
    conn=db()
    try:
        args=request.view_args or {}; unit_id=None; competency_id=args.get("competency_id")
        if endpoint in company_endpoints:
            row=conn.execute("SELECT unit_id FROM companies WHERE id=?",(args.get("company_id"),)).fetchone() if args.get("company_id") else None
            unit_id=row[0] if row else request.form.get("unit_id",type=int)
        elif endpoint in price_endpoints:
            row=conn.execute("SELECT unit_id FROM price_tables WHERE id=?",(args.get("table_id"),)).fetchone() if args.get("table_id") else None
            unit_id=row[0] if row else request.form.get("unit_id",type=int)
        elif endpoint=="unit_delete":unit_id=args.get("unit_id")
        reason=delivery_mutation_reason(conn,unit_id,competency_id,include_uncertain=endpoint!="settings")
    finally:conn.close()
    if reason:
        flash(reason,"error")
        if competency_id:return redirect(url_for("competency_detail",competency_id=competency_id))
        return redirect(url_for("settings" if endpoint in global_endpoints else "registrations"))


@app.route("/setup", methods=["GET", "POST"])
def setup():
    if not local_auth_enabled():
        return redirect(url_for("dashboard"))
    if setting_get("admin_password_hash"):
        return redirect(url_for("login"))
    if request.method == "POST":
        password = request.form.get("password", "")
        confirm = request.form.get("confirm", "")
        if len(password) < 8:
            flash("A senha deve ter pelo menos 8 caracteres.", "error")
        elif password != confirm:
            flash("As senhas não conferem.", "error")
        else:
            setting_set("admin_password_hash", generate_password_hash(password))
            session["authenticated"] = True
            flash("Acesso criado com sucesso.", "success")
            return redirect(url_for("dashboard"))
    return render_template("setup.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if not local_auth_enabled():
        return redirect(url_for("dashboard"))
    if not setting_get("admin_password_hash"):
        return redirect(url_for("setup"))
    if request.method == "POST":
        if check_password_hash(setting_get("admin_password_hash"), request.form.get("password", "")):
            session["authenticated"] = True
            return redirect(url_for("dashboard"))
        flash("Senha incorreta.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def smtp_config():
    security = setting_get("smtp_security", "starttls") or "starttls"
    return {
        "host": setting_get("smtp_host", "smtp.gmail.com") or "smtp.gmail.com",
        "port": 465 if security == "ssl" else 587,
        "username": setting_get("smtp_username").strip(),
        "password": decrypt_secret(setting_get("smtp_password")),
        "security": security,
        "sender_name": setting_get("sender_name", "EDGE Saúde Ocupacional") or "EDGE Saúde Ocupacional",
        "sender_email": (setting_get("sender_email") or setting_get("smtp_username")).strip(),
        "test_mode": setting_get("test_mode", "1") == "1",
        "test_email": setting_get("test_email").strip(),
        "email_signature": setting_get("email_signature", "EDGE Saúde Ocupacional"),
    }


def smtp_send(to_email, cc_value, subject, html_body, text_body, attachments=None, config=None):
    return send_email(config or smtp_config(), valid_email, to_email, cc_value, subject,
                      html_body, text_body, attachments)


@app.after_request
def security_headers(response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    return response


@app.context_processor
def helpers():
    return {
        "csrf_token": generate_csrf_token,
        "money": money,
        "format_cnpj": format_cnpj,
        "format_cpf": format_cpf,
        "format_document": format_document,
        "month_label": month_label,
        "MONTHS": MONTHS,
        "app_version": APP_VERSION,
        "user_role": role(),
        "local_auth_enabled": local_auth_enabled,
        "parent_auth_enabled": parent_auth_enabled,
    }


def get_unit(unit_id):
    conn = db(); row = conn.execute("SELECT * FROM units WHERE id=?", (unit_id,)).fetchone(); conn.close(); return row


def get_competency(comp_id):
    conn = db(); row = conn.execute("SELECT c.*,u.name AS unit_name FROM competencies c JOIN units u ON u.id=c.unit_id WHERE c.id=?", (comp_id,)).fetchone(); conn.close(); return row


def get_company(company_id):
    conn = db(); row = conn.execute("SELECT c.*,u.name AS unit_name FROM companies c JOIN units u ON u.id=c.unit_id WHERE c.id=?", (company_id,)).fetchone(); conn.close(); return row


def sync_competency_companies(conn, competency_id):
    comp = conn.execute("SELECT * FROM competencies WHERE id=?", (competency_id,)).fetchone()
    if not comp or comp["status"] in {"FECHADA"}:
        return
    companies = conn.execute("SELECT * FROM companies WHERE unit_id=? AND active=1 ORDER BY name", (comp["unit_id"],)).fetchall()
    for c in companies:
        conn.execute(
            """INSERT OR IGNORE INTO competency_companies
               (competency_id,company_id,fixed_amount,complementary_amount,created_at,updated_at,fixed_value_pending,complementary_value_pending)
               VALUES(?,?,?,0,?,?,?,?)""",
            (competency_id, c["id"], float(c["fixed_value"] or 0), now_iso(), now_iso(), c["fixed_value_required"], c["complementary_value_required"]),
        )
        # Se ainda não foi enviada, a mensalidade pode acompanhar o valor atual do cadastro.
        if comp["status"] == "ENVIADA":
            # Competências legadas preservam valores existentes, mas aceitam novos cadastros.
            continue
        conn.execute(
            """UPDATE competency_companies SET fixed_amount=?,updated_at=?
               WHERE competency_id=? AND company_id=? AND fixed_sent_at IS NULL AND fixed_paid=0
               AND COALESCE(fixed_amount_manual,0)=0""",
            (float(c["fixed_value"] or 0), now_iso(), competency_id, c["id"]),
        )


def assert_competency_mutable(conn, competency_id, complementary=False):
    """Protege o histórico financeiro e impede alterações durante um envio."""
    comp = conn.execute("SELECT * FROM competencies WHERE id=?", (competency_id,)).fetchone()
    if not comp:
        raise ValueError("Competência não encontrada.")
    if competency_is_closed(comp):
        raise ValueError("A competência está FECHADA. Reabra antes de alterar os dados.")
    if conn.execute("SELECT 1 FROM send_jobs WHERE competency_id=? AND status IN ('QUEUED','RUNNING')", (competency_id,)).fetchone():
        raise ValueError("Há envio em andamento. Aguarde terminar antes de alterar os dados da competência.")
    unresolved = conn.execute(
        """SELECT g.status,g.delivery_result_json FROM send_job_groups g
           JOIN send_jobs j ON j.id=g.job_id WHERE j.competency_id=?
           AND g.status IN ('SENDING','UNCERTAIN','PARTIAL','TEST_PARTIAL')""", (competency_id,)
    ).fetchall()
    for delivery in unresolved:
        try:
            result = json.loads(delivery["delivery_result_json"] or "{}")
        except (TypeError, ValueError):
            result = {}
        if delivery["status"] in {"SENDING", "UNCERTAIN"} or not isinstance(result, dict) or not result.get("primary_accepted"):
            raise ValueError("Existe envio com resultado pendente de conferência. Confira a fila e o histórico antes de alterar os dados da competência.")
    if complementary and conn.execute(
        "SELECT 1 FROM competency_companies WHERE competency_id=? AND (complementary_sent_at IS NOT NULL OR complementary_paid=1)",
        (competency_id,),
    ).fetchone():
        raise ValueError("A apuração possui complementares enviados ou pagos e deve ser preservada. Use uma nova competência para ajustes.")
    return comp


def control_header(ws):
    """Localiza o cabeçalho sem confiar na dimensão residual do Excel."""
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=30, max_col=200, values_only=True), 1):
        mapping = {}
        duplicates = set()
        for column, value in enumerate(row, 1):
            header = norm_header(value)
            if header:
                if header in mapping:
                    duplicates.add(header)
                mapping[header] = column
        employee = any(norm_header(name) in mapping for name in ("FUNCIONARIO", "COLABORADOR", "NOME DO FUNCIONARIO"))
        exam = any(norm_header(name) in mapping for name in ("TIPO DE EXAME", "EXAME"))
        company = any(norm_header(name) in mapping for name in ("SETOR", "EMPRESA", "RAZAO SOCIAL", "CNPJ/CPF", "CNPJ", "CPF DA EMPRESA", "CPF EMPRESA", "CPF"))
        if employee and exam and company:
            if duplicates:
                raise ValueError("O cabeçalho do controle possui colunas repetidas: " + ", ".join(sorted(duplicates)))
            return row_idx, mapping
    raise ValueError("Não encontrei FUNCIONARIO, TIPO DE EXAME e SETOR/EMPRESA ou CNPJ/CPF na planilha de controle.")


def control_sheet_data_bounds(path: Path, sheet_name: str):
    """Retorna (última_linha_com_valor, última_coluna_com_valor) sem confiar na dimensão do Excel.

    Alguns controles reais de setembro possuem mais de um milhão de linhas apenas
    formatadas. O Excel grava essas linhas no XML e o openpyxl pode acabar varrendo
    todas elas mesmo havendo poucos atendimentos. Esta leitura olha somente células
    que efetivamente contêm valor/fórmula/inline string e permite limitar iter_rows.
    """
    relationship_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    spreadsheet_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    with zipfile.ZipFile(path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        relationships = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {
            rel.attrib.get("Id"): rel.attrib.get("Target", "")
            for rel in relationships.findall(f"{{{relationship_ns}}}Relationship")
        }
        target = None
        for sheet in workbook.findall(f".//{{{spreadsheet_ns}}}sheet"):
            if sheet.attrib.get("name") == sheet_name:
                rid = sheet.attrib.get(f"{{{office_rel_ns}}}id")
                target = rel_targets.get(rid)
                break
        if not target:
            return 0, 0
        target = target.replace("\\", "/").lstrip("/")
        member = target if target.startswith("xl/") else f"xl/{target}"
        xml = zf.read(member)

    # Nos controles, células com conteúdo têm <v>, <is> ou <f>. As linhas apenas
    # formatadas possuem células autocontidas (<c .../>) e não entram no limite.
    marker = re.compile(br"<(?:v|is|f)(?:\s[^>]*)?>")
    ref_pattern = re.compile(br'\br="([A-Z]{1,3})(\d+)"')
    last_row = 0
    last_col = 0
    for found in marker.finditer(xml):
        start = xml.rfind(b"<c ", max(0, found.start() - 4096), found.start())
        if start < 0:
            continue
        close = xml.find(b">", start, found.start() + 1)
        if close < 0 or close > found.start():
            continue
        ref = ref_pattern.search(xml[start:close + 1])
        if not ref:
            continue
        letters = ref.group(1).decode("ascii")
        row_number = int(ref.group(2))
        column_number = 0
        for char in letters:
            column_number = column_number * 26 + (ord(char) - 64)
        last_row = max(last_row, row_number)
        last_col = max(last_col, column_number)
    return last_row, last_col


def control_company_documents(value):
    """Extrai identificadores completos da empresa, sem usar CPF do paciente."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        document = _excel_company_document(value)
        return [document] if len(document) in {11, 14} else []
    text = str(value or "")
    pattern = r"(?<!\d)(?:\d{2}[. ]?\d{3}[. ]?\d{3}[/ ]?\d{4}[- ]?\d{2}|\d{3}[. ]?\d{3}[. ]?\d{3}[- ]?\d{2})(?!\d)"
    return list(dict.fromkeys(digits(match.group()) for match in re.finditer(pattern, text)))


def cnpj_from_company_cell(value):
    documents = control_company_documents(value)
    return documents[0] if len(documents) == 1 else ""


def company_name_from_cell(value):
    text = str(value or "").strip()
    pattern = r"(?<!\d)(?:\d{2}[. ]?\d{3}[. ]?\d{3}[/ ]?\d{4}[- ]?\d{2}|\d{3}[. ]?\d{3}[. ]?\d{3}[- ]?\d{2})(?!\d)"
    return re.sub(r"\s+", " ", re.sub(pattern, "", text)).strip(" /-|")


def row_cancelled(row_values):
    joined = " | ".join(normalize_text(x) for x in row_values if x not in (None, ""))
    return any(normalize_text(t) in joined for t in CANCEL_TERMS)




def validate_xlsx_archive(path: Path):
    """Evita planilhas XLSX malformadas/compactadas de forma anormal derrubarem o processo."""
    try:
        with zipfile.ZipFile(path) as zf:
            members = zf.infolist()
            if not {"[Content_Types].xml", "xl/workbook.xml"}.issubset(zf.namelist()):
                raise ValueError("O arquivo não contém uma planilha XLSX válida.")
            if len(members) > 5000:
                raise ValueError("A planilha possui conteúdo interno excessivo.")
            total = sum(max(0, m.file_size) for m in members)
            if total > 300 * 1024 * 1024:
                raise ValueError("A planilha descompactada é grande demais para processamento seguro.")
    except zipfile.BadZipFile:
        raise ValueError("O arquivo XLSX está inválido ou corrompido.")

def process_control_file(competency_id, path: Path, original_name: str, price_source="planilha", auto_register_missing=False):
    """Substitui a apuração atomicamente; DATA não filtra a competência."""
    if price_source != "planilha":
        raise ValueError("Os complementares usam exclusivamente a coluna VALOR da planilha de controle.")
    path = Path(path)
    validate_xlsx_archive(path)
    wb = None
    conn = None
    stored_path = CONTROL_DIR / f"competencia_{competency_id}_{hashlib.sha256(path.read_bytes()).hexdigest()[:12]}.xlsx"
    stored_existed = stored_path.exists()
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        if not wb.worksheets:
            raise ValueError("A planilha não possui uma aba válida.")
        conn = db()
        conn.execute("BEGIN IMMEDIATE")
        comp = assert_competency_mutable(conn, competency_id)
        control_sheets = []
        for candidate in wb.worksheets:
            try:
                control_header(candidate)
            except ValueError:
                continue
            control_sheets.append(candidate)
        if not control_sheets:
            raise ValueError("Nenhuma aba contém o cabeçalho de uma planilha de controle.")
        if len(control_sheets) > 1:
            raise ValueError("Mais de uma aba contém controle de atendimentos. Envie um arquivo com uma única aba de controle para evitar duplicidades.")
        stats, unpriced_names = _process_control_workbook(conn, comp, control_sheets[0], path, original_name, auto_register_missing)
        conn.commit()
    except Exception:
        if conn is not None:
            conn.rollback()
        if not stored_existed and path.resolve() != stored_path.resolve():
            stored_path.unlink(missing_ok=True)
        raise
    finally:
        if conn is not None:
            conn.close()
        if wb is not None:
            wb.close()
    recompute_competency_status(competency_id)
    audit_event("PROCESSAR_CONTROLE", "competencia", competency_id, f"Planilha {original_name} processada.", stats, competency_id=competency_id)
    return stats, unpriced_names


def _process_control_workbook(conn, comp, ws, path, original_name, auto_register_missing=False):
    return control_apuration.process(sys.modules[__name__], conn, comp, ws, path, original_name, auto_register_missing)


def safe_zip_members(zf):
    members = [m for m in zf.infolist() if not m.is_dir()]
    if len(members) > 2500:
        raise ValueError("O ZIP possui arquivos demais. Limite: 2.500 arquivos.")
    total = sum(max(0, m.file_size) for m in members)
    if total > 600 * 1024 * 1024:
        raise ValueError("O conteúdo descompactado do ZIP é muito grande.")
    return members


def _document_company_from_name(filename, companies):
    """Identifica a empresa pelo CNPJ OU CPF cadastrado no campo principal.

    O identificador pode aparecer em qualquer parte do nome do arquivo, com ou sem
    pontuação. CNPJs (14 dígitos) têm prioridade sobre CPFs (11 dígitos).
    """
    numeric_name = digits(Path(filename).stem)
    if not numeric_name:
        return None, "SEM CNPJ/CPF NO NOME"

    cnpj_matches = []
    cpf_matches = []
    for company in companies:
        doc = digits(company["cnpj"])
        if len(doc) == 14 and doc in numeric_name:
            cnpj_matches.append(company)
        elif len(doc) == 11 and doc in numeric_name:
            cpf_matches.append(company)

    unique_cnpj = {c["id"]: c for c in cnpj_matches}
    if len(unique_cnpj) == 1:
        return next(iter(unique_cnpj.values())), None
    if len(unique_cnpj) > 1:
        return None, "MAIS DE UM CNPJ IDENTIFICADO"

    unique_cpf = {c["id"]: c for c in cpf_matches}
    if len(unique_cpf) == 1:
        return next(iter(unique_cpf.values())), None
    if len(unique_cpf) > 1:
        return None, "CPF VINCULADO A MAIS DE UMA EMPRESA"
    return None, "CNPJ/CPF NÃO CADASTRADO"


def accepts_complementary_documents(conn, company):
    """A responsável pode anexar o boleto agrupado mesmo sem cobrança própria."""
    if company["bill_complementaries"]:
        return True
    return bool(conn.execute(
        "SELECT 1 FROM billing_groups WHERE primary_company_id=? AND unit_id=? AND consolidate_complementary=1",
        (company["id"], company["unit_id"]),
    ).fetchone())


def import_documents_zip(competency_id, billing_type, file_storage):
    """Importa o lote inteiro ou remove todas as alterações em caso de falha."""
    if billing_type not in {"FIXED", "COMPLEMENTARY"}:
        raise ValueError("Tipo de cobrança inválido.")
    raw = file_storage.read()
    if not raw:
        raise ValueError("ZIP vazio.")
    conn = None
    written_paths = []
    added = duplicates = unmatched_count = skipped = 0
    unmatched = []
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            members = safe_zip_members(zf)
            if not members:
                raise ValueError("O ZIP não contém arquivos.")
            if zf.testzip() is not None:
                raise ValueError("O ZIP possui um arquivo corrompido. Nenhum documento foi importado.")
            conn = db()
            conn.execute("BEGIN IMMEDIATE")
            assert_competency_mutable(conn, competency_id)
            sync_competency_companies(conn, competency_id)
            companies = conn.execute(
                """SELECT c.* FROM competency_companies cc JOIN companies c ON c.id=cc.company_id
                   WHERE cc.competency_id=?""", (competency_id,)
            ).fetchall()
            run_id = conn.execute(
                """INSERT INTO document_import_runs(competency_id,billing_type,zip_name,total_files,created_at)
                   VALUES(?,?,?,?,?)""",
                (competency_id, billing_type, Path(file_storage.filename or "documentos.zip").name, len(members), now_iso()),
            ).lastrowid
            for m in members:
                name = Path(m.filename).name
                ext = Path(name).suffix.lower()
                result = "IGNORADO"; reason = ""; company = None
                if ext not in ALLOWED_DOC_EXT:
                    skipped += 1; reason = "FORMATO NÃO ACEITO"
                elif m.file_size > 30 * 1024 * 1024:
                    skipped += 1; reason = "ARQUIVO MAIOR QUE 30 MB"
                elif m.file_size == 0:
                    skipped += 1; reason = "ARQUIVO VAZIO"
                else:
                    company, reason = _document_company_from_name(name, companies)
                    if not company:
                        unmatched_count += 1
                        unmatched.append(f"{name} — {reason}")
                        result = "NÃO IDENTIFICADO"
                    elif billing_type == "COMPLEMENTARY" and not accepts_complementary_documents(conn, company):
                        skipped += 1; reason = "EMPRESA PAGA COMPLEMENTARES NO ATO"
                    else:
                        data = zf.read(m)
                        sha = hashlib.sha256(data).hexdigest()
                        exists = conn.execute(
                            "SELECT id FROM documents WHERE competency_id=? AND company_id=? AND billing_type=? AND sha256=?",
                            (competency_id, company["id"], billing_type, sha),
                        ).fetchone()
                        if exists:
                            duplicates += 1; result = "DUPLICADO"; reason = "MESMO ARQUIVO JÁ ANEXADO"
                        else:
                            out_dir = DOCS_DIR / str(competency_id) / billing_type.lower() / str(company["id"])
                            out_dir.mkdir(parents=True, exist_ok=True)
                            out_path = out_dir / f"{uuid.uuid4().hex}{ext}"
                            conn.execute(
                                """INSERT INTO documents(competency_id,company_id,billing_type,original_name,stored_name,sha256,created_at)
                                   VALUES(?,?,?,?,?,?,?)""",
                                (competency_id, company["id"], billing_type, name, str(out_path.relative_to(DATA_DIR)), sha, now_iso()),
                            )
                            written_paths.append(out_path)
                            out_path.write_bytes(data)
                            added += 1; result = "IMPORTADO"; reason = "ASSOCIADO AUTOMATICAMENTE"
                conn.execute(
                    """INSERT INTO document_import_items(run_id,filename,company_id,company_name,result,reason,created_at)
                       VALUES(?,?,?,?,?,?,?)""",
                    (run_id, name, company["id"] if company else None, company["name"] if company else "", result, reason or "", now_iso()),
                )
            conn.execute(
                "UPDATE document_import_runs SET added=?,duplicates=?,unmatched=?,skipped=? WHERE id=?",
                (added, duplicates, unmatched_count, skipped, run_id),
            )
            conn.commit()
    except Exception as exc:
        if conn is not None:
            conn.rollback()
        for out_path in written_paths:
            out_path.unlink(missing_ok=True)
        if isinstance(exc, zipfile.BadZipFile):
            raise ValueError("O arquivo ZIP está inválido ou corrompido. Nenhum documento foi importado.") from exc
        raise
    finally:
        if conn is not None:
            conn.close()
    recompute_competency_status(competency_id)
    audit_event("IMPORTAR_DOCUMENTOS", "competencia", competency_id,
                f"ZIP {billing_type}: {added} importados, {duplicates} duplicados, {unmatched_count} não identificados.",
                {"run_id":run_id,"added":added,"duplicates":duplicates,"unmatched":unmatched_count,"skipped":skipped}, competency_id=competency_id)
    return added, skipped + duplicates, unmatched, run_id, {"duplicates":duplicates,"unmatched":unmatched_count,"skipped":skipped,"total":len(members)}


def company_documents(conn, competency_id, company_id, billing_type):
    rows = conn.execute(
        "SELECT * FROM documents WHERE competency_id=? AND company_id=? AND billing_type=? ORDER BY id",
        (competency_id, company_id, billing_type),
    ).fetchall()
    out = []
    for r in rows:
        p = stored_file_path(r["stored_name"])
        if p.exists():
            out.append({"path": str(p), "filename": r["original_name"]})
    return out


def complementary_summary(conn, competency_id, company_id):
    rows = conn.execute(
        """SELECT exam_name,COUNT(*) qty,MAX(unit_price) unit_price,SUM(COALESCE(total,0)) total
           FROM exam_items WHERE competency_id=? AND company_id=? AND status='OK'
           GROUP BY exam_key,unit_price ORDER BY exam_name,unit_price""",
        (competency_id, company_id),
    ).fetchall()
    return rows


def _build_individual_email(conn, competency_id, company_id, email_type):
    row = conn.execute(
        """SELECT cc.*,c.name,c.cnpj,c.email,c.email_cc,c.bill_complementaries,cp.month,cp.year,u.name unit_name
           FROM competency_companies cc JOIN companies c ON c.id=cc.company_id
           JOIN competencies cp ON cp.id=cc.competency_id JOIN units u ON u.id=cp.unit_id
           WHERE cc.competency_id=? AND cc.company_id=?""",
        (competency_id, company_id),
    ).fetchone()
    if not row:
        return None
    comp_label = month_label(row["month"], row["year"])
    signature=setting_get("email_signature","EDGE Saúde Ocupacional")
    is_reminder = email_type.endswith("_REMINDER")
    billing = "FIXED" if email_type.startswith("FIXED") else "COMPLEMENTARY"
    if billing == "FIXED":
        amount = float(row["fixed_amount"] or 0)
        title = "MENSALIDADE"
        if is_reminder:
            subject = f"LEMBRETE DE PAGAMENTO - MENSALIDADE - {comp_label} - {row['name']}"
            intro = f"Identificamos que a mensalidade referente à competência <strong>{comp_label}</strong> ainda consta como pendente em nosso controle."
        else:
            subject = f"MENSALIDADE - {comp_label} - {row['name']}"
            intro = f"Encaminhamos os documentos referentes à <strong>mensalidade dos serviços de Saúde e Segurança do Trabalho</strong>, competência <strong>{comp_label}</strong>."
        details_html = f"<div style='font-size:18px;font-weight:800;margin:16px 0'>Valor da mensalidade: {html.escape(money(amount))}</div>"
        details_text = f"Valor da mensalidade: {money(amount)}"
    else:
        amount = float(row["complementary_amount"] or 0)
        title = "EXAMES COMPLEMENTARES"
        if is_reminder:
            subject = f"LEMBRETE DE PAGAMENTO - EXAMES COMPLEMENTARES - {comp_label} - {row['name']}"
            intro = f"Identificamos que a cobrança dos exames complementares da competência <strong>{comp_label}</strong> ainda consta como pendente em nosso controle."
        else:
            subject = f"EXAMES COMPLEMENTARES - {comp_label} - {row['name']}"
            intro = f"Encaminhamos a cobrança referente aos <strong>exames complementares realizados</strong> na competência <strong>{comp_label}</strong>."
        summary = complementary_summary(conn, competency_id, company_id)
        tr = "".join(
            f"<tr><td style='padding:7px;border-bottom:1px solid #e5e7eb'>{html.escape(r['exam_name'])}</td>"
            f"<td style='padding:7px;border-bottom:1px solid #e5e7eb;text-align:center'>{r['qty']}</td>"
            f"<td style='padding:7px;border-bottom:1px solid #e5e7eb;text-align:right'>{html.escape(money(r['unit_price']))}</td>"
            f"<td style='padding:7px;border-bottom:1px solid #e5e7eb;text-align:right'>{html.escape(money(r['total']))}</td></tr>"
            for r in summary
        )
        details_html = (
            "<table style='width:100%;border-collapse:collapse;margin:16px 0'>"
            "<thead><tr style='background:#f3f6f8'><th style='padding:8px;text-align:left'>Exame</th><th>Qtd.</th><th style='text-align:right'>Valor unit.</th><th style='text-align:right'>Total</th></tr></thead>"
            f"<tbody>{tr}</tbody></table><div style='font-size:18px;font-weight:800;margin:16px 0'>Total da cobrança: {html.escape(money(amount))}</div>"
        )
        details_text = "\n".join([f"{r['exam_name']}: {r['qty']} x {money(r['unit_price'])} = {money(r['total'])}" for r in summary]) + f"\nTotal: {money(amount)}"
    html_body = (
        "<div style='font-family:Arial,sans-serif;color:#243447;line-height:1.5'>"
        f"<p>Prezados,</p><p><strong>Empresa: {html.escape(row['name'])}</strong><br>CNPJ/CPF: {html.escape(format_document(row['cnpj']))}</p>"
        f"<p>{intro}</p>{details_html}<p>Seguem anexos os respectivos documentos para pagamento.</p>"
        f"<p>Atenciosamente,<br><strong>{html.escape(signature).replace(chr(10),'<br>')}</strong></p></div>"
    )
    text_body = f"Prezados,\n\nEmpresa: {row['name']}\nCNPJ/CPF: {format_document(row['cnpj'])}\n\n{re.sub('<[^>]+>', '', intro)}\n\n{details_text}\n\nSeguem anexos os respectivos documentos para pagamento.\n\n{signature}"
    attachments = company_documents(conn, competency_id, company_id, billing)
    return {"row": row, "subject": subject, "html": html_body, "text": text_body, "attachments": attachments, "amount": amount, "billing": billing, "title": title}


def billing_delivery_ids(conn, competency_id, email_type, company_id, force=False):
    if email_type == 'COMBINED':
        ctx = combined_billing.context(conn, competency_id, company_id)
        return ctx['ids'] if ctx else []
    return delivery_batches.delivery_ids(conn, competency_id, company_id, email_type, force)


def billing_delivery_representative(conn, competency_id, email_type, company_id, force=False):
    if email_type == 'COMBINED':
        ctx = combined_billing.context(conn, competency_id, company_id)
        return ctx['owner'] if ctx else company_id
    return delivery_batches.representative_id(conn, competency_id, company_id, email_type, force)


def billing_send_reason(conn, competency_id, email_type, company_id, force=False):
    if email_type == 'COMBINED':
        return combined_billing.send_reason(sys.modules[__name__], conn, competency_id, company_id)
    return delivery_batches.send_reason(sys.modules[__name__], conn, competency_id, company_id, email_type, force)


def _eligible_billing_company_ids(conn, competency_id, email_type):
    if email_type == 'COMBINED':
        return combined_billing.eligible_ids(sys.modules[__name__], conn, competency_id)
    return delivery_batches.eligible_ids(sys.modules[__name__], conn, competency_id, email_type)


def build_email(conn, competency_id, company_id, email_type, company_ids=None):
    if email_type == 'COMBINED':
        payload = combined_billing.build_payload(sys.modules[__name__], conn, competency_id, company_id, company_ids)
    else:
        payload = delivery_batches.build_payload(sys.modules[__name__], conn, competency_id, company_id, email_type, company_ids)
    payload = billing_policy.enrich_email(sys.modules[__name__], conn, competency_id, payload)
    return complementary_report.enrich(sys.modules[__name__], conn, competency_id, payload)


_send_queue = SendQueue(globals())


def eligible_company_ids(conn, competency_id, email_type):
    return _send_queue.eligible(conn, competency_id, email_type)


def run_send_job(job_id):
    return _send_queue.run(job_id)


def create_send_job_for_ids(competency_id, email_type, ids, parent_job_id=None, force=False):
    return _send_queue.create(competency_id, email_type, ids, parent_job_id=parent_job_id, force=force)


def create_send_job(competency_id, email_type):
    conn = db()
    try:
        ids = eligible_company_ids(conn, competency_id, email_type)
    finally:
        conn.close()
    return create_send_job_for_ids(competency_id, email_type, ids)


def retry_failed_send_job(job_id):
    return _send_queue.retry(job_id)


def recover_send_jobs():
    return _send_queue.recover()


def competency_operational_snapshot(competency_id, selected_filter="all", q=""):
    comp = get_competency(competency_id)
    if not comp:
        return None
    selected_filter = str(selected_filter or "all").strip().lower()
    if selected_filter not in {"all", "blocking", "ready", "payment", "complete", "no_email", "no_docs", "unpriced", "duplicates", "no_value"}:
        selected_filter = "all"
    q = str(q or "").strip()
    q_norm = normalize_text(q)
    conn = db()
    rows = conn.execute(
        """SELECT cc.*,c.id company_id,c.name,c.cnpj,c.email,c.email_cc,c.active,c.bill_complementaries,c.price_table_id,
          (SELECT COUNT(*) FROM exam_items e WHERE e.competency_id=cc.competency_id AND e.company_id=cc.company_id AND e.status='OK') exam_count,
          (SELECT COUNT(*) FROM exam_items e WHERE e.competency_id=cc.competency_id AND e.company_id=cc.company_id AND e.status='SEM_PRECO') unpriced_count,
          (SELECT COUNT(*) FROM exam_items e WHERE e.competency_id=cc.competency_id AND e.company_id=cc.company_id AND e.status='JA_COBRADO') duplicate_count,
          (SELECT COUNT(*) FROM documents d WHERE d.competency_id=cc.competency_id AND d.company_id=cc.company_id AND d.billing_type='FIXED') fixed_docs,
          (SELECT COUNT(*) FROM documents d WHERE d.competency_id=cc.competency_id AND d.company_id=cc.company_id AND d.billing_type='COMPLEMENTARY') comp_docs
          FROM competency_companies cc JOIN companies c ON c.id=cc.company_id
          WHERE cc.competency_id=? ORDER BY c.name""",
        (competency_id,),
    ).fetchall()
    active_job = conn.execute(
        "SELECT * FROM send_jobs WHERE competency_id=? AND status IN ('QUEUED','RUNNING') ORDER BY created_at DESC LIMIT 1",
        (competency_id,),
    ).fetchone()
    totals = conn.execute(
        """SELECT COALESCE(SUM(fixed_amount),0) fixed_total,
        COALESCE(SUM(complementary_amount),0) comp_total,
        COALESCE(SUM(CASE WHEN fixed_paid=1 THEN fixed_amount ELSE 0 END),0) fixed_paid_total,
        COALESCE(SUM(CASE WHEN complementary_paid=1 THEN complementary_amount ELSE 0 END),0) comp_paid_total
        FROM competency_companies WHERE competency_id=?""", (competency_id,)
    ).fetchone()
    # Pendências dos participantes usam os documentos/e-mail da responsável em
    # cada modalidade. O valor e a baixa de cada empresa continuam individuais.
    effective_rows = []
    owner_cache = {}
    for row in rows:
        item = dict(row)
        for email_type, prefix in (("FIXED", "fixed"), ("COMPLEMENTARY", "comp")):
            group_cfg, owner_id, group_rows, billing = grouped_billing.context(conn, competency_id, row["company_id"], email_type)
            item[f"{prefix}_group_name"] = group_cfg["name"] if group_cfg else None
            item[f"{prefix}_owner_id"] = owner_id
            if owner_id not in owner_cache:
                owner_cache[owner_id] = conn.execute("SELECT * FROM companies WHERE id=?", (owner_id,)).fetchone()
            owner = owner_cache[owner_id]
            item[f"{prefix}_email"] = owner["email"] if owner else ""
            item[f"{prefix}_owner_active"] = bool(owner and owner["active"])
            if group_cfg:
                docs = conn.execute("SELECT stored_name FROM documents WHERE competency_id=? AND company_id=? AND billing_type=?", (competency_id,owner_id,billing)).fetchall()
                item[f"{prefix}_docs"] = len(docs) if all(stored_file_path(d["stored_name"]).is_file() for d in docs) else 0
                if email_type == "COMPLEMENTARY":
                    outstanding=grouped_billing.outstanding_rows(conn,competency_id,row["company_id"],"COMPLEMENTARY")[2]
                    item["group_unpriced_count"] = sum(r["unpriced"] for r in outstanding if not r["complementary_amount_manual"])
            else:
                docs = conn.execute("SELECT stored_name FROM documents WHERE competency_id=? AND company_id=? AND billing_type=?", (competency_id,owner_id,billing)).fetchall()
                if any(not stored_file_path(d["stored_name"]).is_file() for d in docs):
                    item[f"{prefix}_docs"] = 0
        effective_rows.append(item)
    rows = effective_rows
    dispatch_counts = {email_type: len(eligible_company_ids(conn,competency_id,email_type)) for email_type in grouped_billing.EMAIL_TYPES}
    delivery_reasons={}
    reason_cache={}
    registered_members={item[0] for item in conn.execute("SELECT company_id FROM billing_group_members")}
    for row in rows:
        for email_type,prefix,docs_key in (("FIXED","fixed","fixed_docs"),("COMPLEMENTARY","complementary","comp_docs")):
            if not row["active"] or row[f"{prefix}_paid"] or float(row[f"{prefix}_amount"] or 0)<=0 or not row[docs_key]:
                continue
            if email_type=="COMPLEMENTARY" and (not row["bill_complementaries"] or (row["unpriced_count"] and not row["complementary_amount_manual"])):
                continue
            kind=email_type+"_REMINDER" if row[f"{prefix}_sent_at"] else email_type
            display_prefix="fixed" if email_type=="FIXED" else "comp"
            if row.get(f"{display_prefix}_group_name"):
                packet_key=("group",row[f"{display_prefix}_owner_id"],kind)
            elif row["company_id"] in registered_members:
                packet_key=("individual",row["company_id"],kind)
            else:
                packet_key=("email",delivery_batches.normalize_email(row.get("email")),kind)
            if packet_key not in reason_cache:
                reason_cache[packet_key]=billing_send_reason(conn,competency_id,kind,row["company_id"])
            delivery_reasons[row["company_id"],email_type]=reason_cache[packet_key]
    for envelope in combined_billing.envelopes(conn,competency_id):
        reason=combined_billing.reason(sys.modules[__name__],conn,competency_id,envelope)
        if reason:
            for packet in envelope['packets']:
                for cid in packet['ids']: delivery_reasons[cid,packet['billing']]=reason
    conn.close()
    control_processed = bool(comp["control_stored_name"] or comp["processed_at"])
    cfg = smtp_config()
    gmail_ready = bool(cfg.get("host") and cfg.get("username") and cfg.get("password") and cfg.get("sender_email"))
    test_mode_warning = bool(cfg.get("test_mode") and not valid_email(cfg.get("test_email")))

    def state(label, kind="neutral", code=""):
        return {"label": label, "kind": kind, "code": code or normalize_text(label).replace(" ", "_")}

    items=[]
    all_items=[]
    counts={"total":0,"blocking":0,"ready":0,"payment":0,"complete":0,"no_email":0,"no_docs":0,"unpriced":0,"no_value":0,"duplicates":0}
    fixed_ready=comp_ready=fixed_waiting=comp_waiting=0
    for row in rows:
        r=dict(row)
        fixed_email_ok=valid_email(r.get("fixed_email")); comp_email_ok=valid_email(r.get("comp_email")); is_active=bool(r.get("active")); fixed_docs=int(r.get("fixed_docs") or 0); comp_docs=int(r.get("comp_docs") or 0)
        unpriced=int(r.get("group_unpriced_count",r.get("unpriced_count")) or 0); duplicate_count=int(r.get("duplicate_count") or 0); fixed_amount=float(r.get("fixed_amount") or 0); comp_amount=float(r.get("complementary_amount") or 0); bill_comp=bool(r.get("bill_complementaries"))
        blockers=[]; blocker_codes=set()
        if duplicate_count: counts["duplicates"] += 1
        if not is_active:
            fixed_state=state("EMPRESA INATIVA","error","INACTIVE"); blockers.append("Empresa inativa"); blocker_codes.add("inactive")
        elif r.get("fixed_paid"):
            fixed_state=state("PAGO","ok","PAID")
        elif r.get("fixed_value_pending") and not r.get("fixed_sent_at"):
            fixed_state=state("CONFERIR VALOR","warn","VALUE_PENDING"); blockers.append("Conferir e confirmar mensalidade na prévia"); blocker_codes.add("no_value")
        elif fixed_amount<=0 and (r.get("fixed_group_name") or r.get("fixed_amount_manual")):
            fixed_state=state("SEM COBRANÇA","neutral","NO_CHARGE")
        elif not r.get("fixed_owner_active"):
            fixed_state=state("RESPONSÁVEL INATIVA","error","INACTIVE"); blockers.append("Empresa responsável pela mensalidade está inativa"); blocker_codes.add("inactive")
        elif not fixed_email_ok:
            fixed_state=state("SEM E-MAIL","error","NO_EMAIL"); blockers.append("E-mail principal ausente ou inválido"); blocker_codes.add("no_email")
        elif fixed_amount<=0:
            fixed_state=state("SEM VALOR","error","NO_VALUE"); blockers.append("Valor da mensalidade não cadastrado"); blocker_codes.add("no_value")
        elif r.get("fixed_sent_at"):
            fixed_state=state("AGUARDANDO PAGAMENTO","warn","PAYMENT") if fixed_docs>0 else state("SEM DOCUMENTO","error","NO_DOC")
            if fixed_docs<=0: blockers.append("Documento da mensalidade ausente para eventual lembrete"); blocker_codes.add("no_docs")
        elif fixed_docs<=0:
            fixed_state=state("SEM DOCUMENTO","error","NO_DOC"); blockers.append("Documento da mensalidade não importado"); blocker_codes.add("no_docs")
        else:
            fixed_state=state("PRONTO PARA ENVIAR","ok","READY")
        if not bill_comp:
            comp_state=state("PAGO NO ATO","neutral","PAID_AT_EXAM")
        elif not is_active:
            comp_state=state("EMPRESA INATIVA","error","INACTIVE"); blocker_codes.add("inactive")
            if "Empresa inativa" not in blockers: blockers.append("Empresa inativa")
        elif r.get("complementary_value_pending") and not r.get("complementary_sent_at"):
            comp_state=state("CONFERIR VALOR","warn","VALUE_PENDING"); blockers.append("Conferir e confirmar complementares na prévia"); blocker_codes.add("no_value")
        elif not control_processed and not r.get("complementary_amount_manual"):
            comp_state=state("AGUARDA CONTROLE","warn","WAIT_CONTROL")
        elif r.get("complementary_paid"):
            comp_state=state("PAGO","ok","PAID")
        elif unpriced>0 and not r.get("complementary_amount_manual"):
            comp_state=state("SEM PREÇO","error","UNPRICED"); blockers.append(f"{unpriced} exame(s) sem preço"); blocker_codes.add("unpriced")
        elif comp_amount<=0:
            comp_state=state("SEM COBRANÇA","neutral","NO_CHARGE")
        elif not r.get("comp_owner_active"):
            comp_state=state("RESPONSÁVEL INATIVA","error","INACTIVE"); blockers.append("Empresa responsável pelos complementares está inativa"); blocker_codes.add("inactive")
        elif not comp_email_ok:
            comp_state=state("SEM E-MAIL","error","NO_EMAIL"); blocker_codes.add("no_email")
            if "E-mail principal ausente ou inválido" not in blockers: blockers.append("E-mail principal ausente ou inválido")
        elif r.get("complementary_sent_at"):
            comp_state=state("AGUARDANDO PAGAMENTO","warn","PAYMENT") if comp_docs>0 else state("SEM DOCUMENTO","error","NO_DOC")
            if comp_docs<=0: blockers.append("Documento dos complementares ausente para eventual lembrete"); blocker_codes.add("no_docs")
        elif comp_docs<=0:
            comp_state=state("SEM DOCUMENTO","error","NO_DOC"); blockers.append("Documento dos complementares não importado"); blocker_codes.add("no_docs")
        else:
            comp_state=state("PRONTO PARA ENVIAR","ok","READY")
        for email_type,current in (("FIXED",fixed_state),("COMPLEMENTARY",comp_state)):
            reason=delivery_reasons.get((r["company_id"],email_type))
            if reason and current["code"] in {"READY","PAYMENT"}:
                replacement=state("ENVIO BLOQUEADO","error","DELIVERY_BLOCKED")
                if email_type=="FIXED":fixed_state=replacement
                else:comp_state=replacement
                blockers.append(("Mensalidade: " if email_type=="FIXED" else "Complementares: ")+reason)
                blocker_codes.add("delivery")
        ready=fixed_state["code"]=="READY" or comp_state["code"]=="READY"
        payment=fixed_state["code"]=="PAYMENT" or comp_state["code"]=="PAYMENT"
        complete=fixed_state["code"] in {"PAID","NO_CHARGE"} and comp_state["code"] in {"PAID","PAID_AT_EXAM","NO_CHARGE"}
        if blockers:
            overall=state("PENDÊNCIA","error","BLOCKING"); group="blocking"
        elif ready:
            overall=state("PRONTO PARA ENVIAR","ok","READY"); group="ready"
        elif payment:
            overall=state("AGUARDANDO PAGAMENTO","warn","PAYMENT"); group="payment"
        elif complete:
            overall=state("CONCLUÍDO","ok","COMPLETE"); group="complete"
        elif comp_state["code"]=="WAIT_CONTROL":
            overall=state("AGUARDA CONTROLE","warn","WAIT_CONTROL"); group="blocking"
        else:
            overall=state("ACOMPANHAR","neutral","OTHER"); group="all"
        r.update({"fixed_state":fixed_state,"comp_state":comp_state,"overall_state":overall,"blockers":blockers,"group":group})
        all_items.append(r); counts["total"]+=1
        if blockers or group=="blocking": counts["blocking"]+=1
        if ready: counts["ready"]+=1
        if payment: counts["payment"]+=1
        if complete: counts["complete"]+=1
        for code in blocker_codes:
            if code in counts: counts[code]+=1
        if fixed_state["code"]=="READY": fixed_ready+=1
        if comp_state["code"]=="READY": comp_ready+=1
        if fixed_state["code"]=="PAYMENT": fixed_waiting+=1
        if comp_state["code"]=="PAYMENT": comp_waiting+=1
        if q_norm:
            haystack=normalize_text(f"{r.get('name','')} {r.get('cnpj','')} {r.get('email','')}")
            if q_norm not in haystack: continue
        if selected_filter=="blocking" and not (blockers or group=="blocking"): continue
        if selected_filter=="ready" and not ready: continue
        if selected_filter=="payment" and not payment: continue
        if selected_filter=="complete" and not complete: continue
        if selected_filter=="no_email" and "no_email" not in blocker_codes: continue
        if selected_filter=="no_docs" and "no_docs" not in blocker_codes: continue
        if selected_filter=="unpriced" and "unpriced" not in blocker_codes: continue
        if selected_filter=="duplicates" and duplicate_count<=0: continue
        if selected_filter=="no_value" and "no_value" not in blocker_codes: continue
        items.append(r)
    fixed_ready = dispatch_counts["FIXED"]
    comp_ready = dispatch_counts["COMPLEMENTARY"]
    fixed_waiting = dispatch_counts["FIXED_REMINDER"]
    comp_waiting = dispatch_counts["COMPLEMENTARY_REMINDER"]
    # Próxima ação operacional sugerida pelo estado real da competência.
    if counts["no_value"]>0:
        next_action={"code":"VALUES","title":"Conferir valores antes do envio","text":f"Há {counts['no_value']} empresa(s) com valores pendentes. Abra a prévia da modalidade para conferir e confirmar.","endpoint":"competency_detail"}
    elif not control_processed:
        next_action={"code":"CONTROL","title":"Processar a planilha de controle","text":"Envie o controle do mês para apurar os exames complementares.","endpoint":"competency_apuration_page"}
    elif counts["unpriced"]>0:
        next_action={"code":"PRICES","title":"Corrigir preços pendentes","text":f"Existem {counts['unpriced']} empresa(s) com exame sem preço.","endpoint":"competency_detail"}
    elif counts["duplicates"]>0:
        next_action={"code":"DUPLICATES","title":"Conferir exames já cobrados","text":f"Existem {counts['duplicates']} empresa(s) com atendimento identificado como já faturado anteriormente. Eles foram bloqueados automaticamente.","endpoint":"competency_apuration_page"}
    elif counts["no_docs"]>0:
        next_action={"code":"DOCS","title":"Importar documentos que faltam","text":f"Existem {counts['no_docs']} empresa(s) aguardando documentos.","endpoint":"competency_documents_page"}
    elif counts["ready"]>0:
        next_action={"code":"SEND","title":"Revisar e enviar cobranças","text":f"Há {counts['ready']} empresa(s) com pelo menos uma cobrança pronta.","endpoint":"competency_review"}
    elif counts["payment"]>0:
        next_action={"code":"PAY","title":"Acompanhar pagamentos","text":f"Há {counts['payment']} empresa(s) aguardando pagamento.","endpoint":"competency_payments_page"}
    else:
        next_action={"code":"DONE","title":"Competência organizada","text":"Não há ação operacional urgente nesta competência.","endpoint":"competency_payments_page"}
    with db() as count_conn:
        missing_companies = count_conn.execute("SELECT COUNT(*) FROM companies c WHERE c.unit_id=? AND c.active=1 AND NOT EXISTS(SELECT 1 FROM competency_companies cc WHERE cc.competency_id=? AND cc.company_id=c.id)",(comp["unit_id"],competency_id)).fetchone()[0]
    count_conn.close()
    return {"missing_companies":missing_companies,"comp":comp,"items":items,"all_items":all_items,"counts":counts,"totals":totals,"control_processed":control_processed,
            "gmail_ready":gmail_ready,"test_mode":bool(cfg.get("test_mode")),"test_email":cfg.get("test_email"),"test_mode_warning":test_mode_warning,
            "active_job":active_job,"selected_filter":selected_filter,"q":q,"next_action":next_action,
            "fixed_ready":fixed_ready,"comp_ready":comp_ready,"fixed_waiting":fixed_waiting,"comp_waiting":comp_waiting}


@app.route("/")
def dashboard():
    conn = db()
    units = conn.execute(
        """SELECT u.*,
          (SELECT COUNT(*) FROM competencies x WHERE x.unit_id=u.id) competencies_count,
          (SELECT COUNT(*) FROM companies c WHERE c.unit_id=u.id AND c.active=1) companies_count,
          (SELECT COALESCE(SUM(cc.fixed_amount+cc.complementary_amount),0) FROM competency_companies cc JOIN competencies x ON x.id=cc.competency_id WHERE x.unit_id=u.id) billed_total,
          (SELECT COUNT(*) FROM competency_companies cc JOIN competencies x ON x.id=cc.competency_id WHERE x.unit_id=u.id AND (cc.fixed_sent_at IS NOT NULL AND cc.fixed_paid=0 OR cc.complementary_sent_at IS NOT NULL AND cc.complementary_amount>0 AND cc.complementary_paid=0)) awaiting_payment
          FROM units u WHERE u.active=1 ORDER BY u.name"""
    ).fetchall()
    recent = conn.execute(
        """SELECT cp.id,cp.month,cp.year,cp.status,u.name unit_name,
        (SELECT COUNT(*) FROM competency_companies cc WHERE cc.competency_id=cp.id) companies_count,
        (SELECT COALESCE(SUM(fixed_amount+complementary_amount),0) FROM competency_companies cc WHERE cc.competency_id=cp.id) total
        FROM competencies cp JOIN units u ON u.id=cp.unit_id
        ORDER BY cp.year DESC,cp.month DESC,cp.id DESC LIMIT 6"""
    ).fetchall()
    global_totals=conn.execute("""SELECT
        COALESCE(SUM(cc.fixed_amount+cc.complementary_amount),0) billed,
        COALESCE(SUM(CASE WHEN cc.fixed_paid=1 THEN cc.fixed_amount ELSE 0 END)+SUM(CASE WHEN cc.complementary_paid=1 THEN cc.complementary_amount ELSE 0 END),0) paid,
        COUNT(DISTINCT CASE WHEN (cc.fixed_sent_at IS NOT NULL AND cc.fixed_paid=0) OR (cc.complementary_sent_at IS NOT NULL AND cc.complementary_amount>0 AND cc.complementary_paid=0) THEN cc.company_id END) awaiting
        FROM competency_companies cc JOIN competencies cp ON cp.id=cc.competency_id WHERE COALESCE(cp.status,'PREPARACAO')!='FECHADA'""").fetchone()
    open_competencies=conn.execute("SELECT id FROM competencies WHERE COALESCE(status,'PREPARACAO')!='FECHADA' ORDER BY year DESC,month DESC LIMIT 24").fetchall()
    conn.close()
    operational_pending=0
    for cp in open_competencies:
        snap=competency_operational_snapshot(cp["id"]); operational_pending += int(snap["counts"]["blocking"] if snap else 0)
    cfg=smtp_config(); gmail_ready=bool(cfg.get("host") and cfg.get("username") and cfg.get("password") and cfg.get("sender_email"))
    global_kpis={"billed":float(global_totals["billed"] or 0),"paid":float(global_totals["paid"] or 0),"pending_value":float(global_totals["billed"] or 0)-float(global_totals["paid"] or 0),"awaiting":int(global_totals["awaiting"] or 0),"operational_pending":operational_pending}
    return render_template("dashboard.html", units=units, recent=recent, gmail_ready=gmail_ready, test_mode=bool(cfg.get("test_mode")),global_kpis=global_kpis)


@app.route("/financial-overview")
def financial_overview():
    selected=str(request.args.get("filter") or "all").lower()
    if selected not in {"all","paid","pending","operational"}: selected="all"
    rows=[]
    if selected=="operational":
        conn=db(); comps=conn.execute("SELECT id FROM competencies WHERE COALESCE(status,'PREPARACAO')!='FECHADA' ORDER BY year DESC,month DESC,id DESC").fetchall(); conn.close()
        for cp in comps:
            snap=competency_operational_snapshot(cp["id"],"blocking")
            if not snap: continue
            for item in snap["items"]:
                rows.append({
                    "competency_id":cp["id"],"unit_name":snap["comp"]["unit_name"],"month":snap["comp"]["month"],"year":snap["comp"]["year"],
                    "company_id":item["company_id"],"name":item["name"],"cnpj":item["cnpj"],"email":item["email"],
                    "fixed_amount":item["fixed_amount"],"complementary_amount":item["complementary_amount"],
                    "fixed_paid":item["fixed_paid"],"complementary_paid":item["complementary_paid"],
                    "fixed_sent_at":item["fixed_sent_at"],"complementary_sent_at":item["complementary_sent_at"],
                    "issues":"; ".join(item["blockers"] or [item["overall_state"]["label"]])
                })
    else:
        conn=db(); raw=conn.execute(
            """SELECT cc.*,c.name,c.cnpj,c.email,cp.month,cp.year,cp.status,u.name unit_name
               FROM competency_companies cc JOIN companies c ON c.id=cc.company_id
               JOIN competencies cp ON cp.id=cc.competency_id JOIN units u ON u.id=cp.unit_id
               WHERE COALESCE(cp.status,'PREPARACAO')!='FECHADA'
               ORDER BY cp.year DESC,cp.month DESC,u.name,c.name"""
        ).fetchall(); conn.close()
        for rr in raw:
            r=dict(rr); total=float(r.get("fixed_amount") or 0)+float(r.get("complementary_amount") or 0)
            paid=float(r.get("fixed_amount") or 0)*(1 if r.get("fixed_paid") else 0)+float(r.get("complementary_amount") or 0)*(1 if r.get("complementary_paid") else 0)
            waiting=(bool(r.get("fixed_sent_at")) and not r.get("fixed_paid")) or (bool(r.get("complementary_sent_at")) and float(r.get("complementary_amount") or 0)>0 and not r.get("complementary_paid"))
            if selected=="paid" and paid<=0: continue
            if selected=="pending" and not waiting: continue
            r["issues"]="Aguardando pagamento" if waiting else ("Pagamento registrado" if paid>0 else "Ainda não enviado")
            rows.append(r)
    totals={"billed":0.0,"paid":0.0,"pending":0.0}
    for r in rows:
        total=float(r.get("fixed_amount") or 0)+float(r.get("complementary_amount") or 0)
        paid=float(r.get("fixed_amount") or 0)*(1 if r.get("fixed_paid") else 0)+float(r.get("complementary_amount") or 0)*(1 if r.get("complementary_paid") else 0)
        totals["billed"]+=total; totals["paid"]+=paid; totals["pending"]+=max(0,total-paid)
    return render_template("financial_overview.html",rows=rows,selected_filter=selected,totals=totals)


@app.route("/units")
def units():
    conn = db(); rows = conn.execute("SELECT * FROM units ORDER BY name").fetchall(); conn.close(); return render_template("units.html", units=rows)


@app.route("/units/new", methods=["GET", "POST"])
def unit_new():
    if request.method == "POST":
        name = str(request.form.get("name") or "").strip().upper()
        if not name:
            flash("Informe o nome da unidade.", "error")
        else:
            try:
                conn = db(); conn.execute("INSERT INTO units(name,active,created_at,updated_at) VALUES(?,1,?,?)", (name, now_iso(), now_iso())); conn.commit(); conn.close(); flash("Unidade cadastrada.", "success"); return redirect(url_for("units"))
            except sqlite3.IntegrityError:
                flash("Essa unidade já existe.", "error")
    return render_template("unit_new.html")


@app.route("/units/<int:unit_id>/delete", methods=["POST"])
def unit_delete(unit_id):
    if not confirm_delete_request():
        flash("Digite EXCLUIR para confirmar a exclusão definitiva da unidade.", "error")
        return redirect(url_for("units"))
    conn=db(); unit=conn.execute("SELECT * FROM units WHERE id=?",(unit_id,)).fetchone()
    if not unit:
        conn.close(); abort(404)
    running=conn.execute("""SELECT 1 FROM send_jobs j JOIN competencies cp ON cp.id=j.competency_id
                            WHERE cp.unit_id=? AND j.status IN ('QUEUED','RUNNING') LIMIT 1""",(unit_id,)).fetchone()
    if running:
        conn.close(); flash("Existe envio em andamento nesta unidade. Aguarde terminar antes de excluir.","error"); return redirect(url_for("units"))
    docs=conn.execute("""SELECT d.stored_name FROM documents d JOIN competencies cp ON cp.id=d.competency_id WHERE cp.unit_id=?""",(unit_id,)).fetchall()
    controls=conn.execute("SELECT control_stored_name FROM competencies WHERE unit_id=? AND control_stored_name IS NOT NULL",(unit_id,)).fetchall()
    try:
        create_db_backup(f"antes_excluir_unidade_{unit_id}")
        conn.execute("BEGIN IMMEDIATE")
        # Competências primeiro: cascatas removem apuração, cobranças, logs, jobs e vínculos.
        conn.execute("DELETE FROM competencies WHERE unit_id=?",(unit_id,))
        # Agora não restam históricos que impeçam apagar as empresas.
        company_ids=[r["id"] for r in conn.execute("SELECT id FROM companies WHERE unit_id=?",(unit_id,)).fetchall()]
        if company_ids:
            q=','.join('?'*len(company_ids))
            conn.execute(f"DELETE FROM send_job_groups WHERE company_id IN ({q})",company_ids)
        conn.execute("DELETE FROM companies WHERE unit_id=?",(unit_id,))
        table_ids=[r["id"] for r in conn.execute("SELECT id FROM price_tables WHERE unit_id=?",(unit_id,)).fetchall()]
        for tid in table_ids:
            conn.execute("DELETE FROM price_table_items WHERE table_id=?",(tid,))
        conn.execute("DELETE FROM price_tables WHERE unit_id=?",(unit_id,))
        conn.execute("DELETE FROM units WHERE id=?",(unit_id,))
        conn.commit()
    except Exception as exc:
        conn.rollback(); conn.close(); flash(f"Não foi possível excluir a unidade: {exc}","error"); return redirect(url_for("units"))
    conn.close(); remove_document_files(docs)
    for r in controls: remove_control_file(r["control_stored_name"])
    flash(f"Unidade {unit['name']} e todos os dados vinculados foram excluídos definitivamente.","success")
    return redirect(url_for("units"))


@app.route("/units/<int:unit_id>")
def unit_dashboard(unit_id):
    unit = get_unit(unit_id)
    if not unit: abort(404)
    conn = db()
    base = conn.execute(
        """SELECT cp.*,
          (SELECT COUNT(*) FROM competency_companies cc WHERE cc.competency_id=cp.id) companies_count,
          (SELECT COALESCE(SUM(cc.fixed_amount),0) FROM competency_companies cc WHERE cc.competency_id=cp.id) fixed_total,
          (SELECT COALESCE(SUM(cc.complementary_amount),0) FROM competency_companies cc WHERE cc.competency_id=cp.id) comp_total
          FROM competencies cp WHERE cp.unit_id=? ORDER BY cp.year DESC,cp.month DESC""", (unit_id,)
    ).fetchall()
    company_count=conn.execute("SELECT COUNT(*) n FROM companies WHERE unit_id=? AND active=1",(unit_id,)).fetchone()["n"]
    conn.close()
    comps=[]
    for row in base:
        d=dict(row); snap=competency_operational_snapshot(row["id"]); d["counts"]=snap["counts"]; d["next_action"]=snap["next_action"]; d["control_processed"]=snap["control_processed"]; d["paid_total"]=float(snap["totals"]["fixed_paid_total"] or 0)+float(snap["totals"]["comp_paid_total"] or 0); comps.append(d)
    return render_template("unit_dashboard.html", unit=unit, competencies=comps, company_count=company_count)



def _excel_sheet_by_names(wb, *names):
    wanted={norm_header(n) for n in names}
    for ws in wb.worksheets:
        if norm_header(ws.title) in wanted:
            return ws
    return None


def _header_index(ws):
    out={}
    for col in range(1, ws.max_column+1):
        key=norm_header(ws.cell(1,col).value)
        if key:
            if key in out:
                raise ValueError(f"A aba {ws.title} possui uma coluna repetida: {ws.cell(1,col).value}.")
            out[key]=col
    return out


def _col_from(headers, *aliases):
    for a in aliases:
        k=norm_header(a)
        if k in headers:
            return headers[k]
    return None


def _truthy_excel(value, default=True):
    if value is None or str(value).strip()=="":
        return default
    s=normalize_text(value)
    if s in {"SIM","S","1","TRUE","ATIVA","ATIVO","YES","FATURADO","FATURAR","FATURADO NO MES","MENSAL","COBRAR"}:
        return True
    if s in {"NAO","N","0","FALSE","INATIVA","INATIVO","NO","PAGO NO ATO","PAGA NO ATO","NA HORA","PAGO NA HORA","PAGA NA HORA"}:
        return False
    return default


def _billing_mode_excel(value):
    s=normalize_text(value)
    if s in {"FATURADO NO MES","FATURADO","FATURAR","MENSAL","SIM","S","COBRAR"}:
        return 1
    if s in {"PAGO NO ATO","PAGA NO ATO","NA HORA","PAGO NA HORA","PAGA NA HORA","NAO","N"}:
        return 0
    return None


def active_exam_types(conn):
    return conn.execute("SELECT * FROM complementary_exam_types WHERE active=1 ORDER BY name").fetchall()


def price_tables_for_unit(conn, unit_id, active_only=True):
    sql="SELECT pt.*, (SELECT COUNT(*) FROM companies c WHERE c.price_table_id=pt.id) company_count FROM price_tables pt WHERE pt.unit_id=?"
    params=[unit_id]
    if active_only:
        sql += " AND pt.active=1"
    sql += " ORDER BY pt.name"
    return conn.execute(sql, params).fetchall()


def apply_price_table_to_company(conn, company_id, table_id):
    company=conn.execute("SELECT * FROM companies WHERE id=?",(company_id,)).fetchone()
    if not company:
        raise ValueError("Empresa não encontrada.")
    if not table_id:
        conn.execute("UPDATE companies SET price_table_id=NULL,updated_at=? WHERE id=?",(now_iso(),company_id))
        return
    table=conn.execute("SELECT * FROM price_tables WHERE id=? AND unit_id=?",(table_id,company['unit_id'])).fetchone()
    if not table:
        raise ValueError("Tabela de preços não pertence à unidade da empresa.")
    items=conn.execute("SELECT * FROM price_table_items WHERE table_id=? ORDER BY exam_name",(table_id,)).fetchall()
    conn.execute("DELETE FROM complementary_prices WHERE company_id=?",(company_id,))
    for item in items:
        conn.execute("INSERT INTO complementary_prices(company_id,exam_name,exam_key,unit_price,created_at,updated_at) VALUES(?,?,?,?,?,?)",
                     (company_id,item['exam_name'],item['exam_key'],item['unit_price'],now_iso(),now_iso()))
    conn.execute("UPDATE companies SET price_table_id=?,updated_at=? WHERE id=?",(table_id,now_iso(),company_id))


def sync_price_table_companies(conn, table_id):
    company_ids=[r['id'] for r in conn.execute("SELECT id FROM companies WHERE price_table_id=? AND bill_complementaries=1",(table_id,)).fetchall()]
    for cid in company_ids:
        apply_price_table_to_company(conn,cid,table_id)
    return len(company_ids)


def _import_workbook(file_storage):
    data = file_storage.read()
    if not data:
        raise ValueError("A planilha está vazia.")
    validate_xlsx_archive(io.BytesIO(data))
    wb = load_workbook(io.BytesIO(data), data_only=True)
    if not wb.worksheets or any(ws.max_row > 100000 or ws.max_column > 200 for ws in wb.worksheets):
        wb.close()
        raise ValueError("A planilha deve possuir abas válidas, até 100.000 linhas e 200 colunas por aba.")
    return wb


def _excel_company_document(value):
    if isinstance(value, bool):
        return ""
    if isinstance(value, (int, float)):
        if not isinstance(value, int) and (not value.is_integer()):
            return ""
        value = str(int(value))
    doc = digits(value)
    # O Excel remove zeros à esquerda quando CNPJ/CPF é colado como número.
    # Recupera apenas os dois casos inequívocos: 10→11 (CPF) e 13→14 (CNPJ).
    if len(doc) == 10:
        doc = doc.zfill(11)
    elif len(doc) == 13:
        doc = doc.zfill(14)
    return doc


def _excel_active(value, default=True):
    if value is None or str(value).strip() == "":
        return int(bool(default))
    key = normalize_text(value)
    if key in {"SIM", "S", "1", "TRUE", "ATIVA", "ATIVO", "YES"}:
        return 1
    if key in {"NAO", "N", "0", "FALSE", "INATIVA", "INATIVO", "NO"}:
        return 0
    raise ValueError(f"ATIVA inválido '{value}'. Use SIM ou NÃO.")


def _import_email_tokens(value):
    raw = str(value or "").strip().lower()
    if not raw:
        return []
    # A base da EDGE usa '/', ';', vírgula, hífen e até apenas espaços entre
    # endereços. Em vez de depender de um único separador, extraímos os e-mails
    # válidos diretamente da célula e validamos se sobrou algum texto estranho.
    # O caractere "/" é tratado como separador de destinatários na base EDGE,
    # por isso não entra na parte local do padrão, mesmo sendo permitido em casos
    # raros pelo padrão geral de e-mail.
    pattern = r"[a-z0-9.!#$%&'*+=?^_`{|}~-]+@[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)+"
    tokens = re.findall(pattern, raw, flags=re.IGNORECASE)
    residue = re.sub(pattern, "", raw, flags=re.IGNORECASE)
    residue = re.sub(r"[\s;,/|:\-]+", "", residue)
    if residue or not tokens:
        raise ValueError(f"e-mail inválido '{str(value).strip()}'.")
    out = []
    for token in tokens:
        token = token.strip().lower()
        if valid_email(token) and token not in out:
            out.append(token)
    if not out:
        raise ValueError(f"e-mail inválido '{str(value).strip()}'.")
    return out


def _merge_email_lists(*values):
    merged = []
    for value in values:
        for email in value or []:
            if email and email not in merged:
                merged.append(email)
    return merged


def import_companies_workbook(file_storage):
    """Importa o cadastro completo com validação transacional e consolidação por CNPJ/CPF.

    Regras V2.10:
    - CNPJ/CPF pode vir formatado ou somente com dígitos;
    - múltiplos e-mails podem vir juntos em EMAIL usando /, ;, vírgula, hífen ou espaços;
    - linhas repetidas do mesmo documento/unidade são consolidadas, somando mensalidades;
    - nenhum dado é gravado quando existe erro real de validação.
    """
    wb = conn = None
    details = []
    try:
        wb = _import_workbook(file_storage)
        ws = _excel_sheet_by_names(wb, "EMPRESAS", "EMPRESA") or wb.worksheets[0]
        h = _header_index(ws)
        c_unit = _col_from(h, "UNIDADE", "FILIAL", "LOCAL")
        c_doc = _col_from(h, "CNPJ/CPF", "CNPJ CPF", "CNPJ", "CPF")
        c_name = _col_from(h, "EMPRESA", "NOME", "RAZAO SOCIAL", "RAZÃO SOCIAL")
        c_email = _col_from(h, "EMAIL", "E-MAIL", "EMAIL PRINCIPAL", "E-MAIL PRINCIPAL")
        c_cc = _col_from(h, "EMAIL_CC", "EMAIL CC", "E-MAIL CC", "CC")
        c_fixed = _col_from(h, "VALOR FIXO MENSAL", "VALOR_FIXO_MENSAL", "VALOR FIXO", "MENSALIDADE", "VALOR MENSAL")
        c_active = _col_from(h, "ATIVA", "ATIVO", "STATUS")
        c_billing = _col_from(h, "COMPLEMENTARES", "COBRANCA_COMPLEMENTARES", "COBRANÇA_COMPLEMENTARES")
        missing = [name for name, col in [("UNIDADE", c_unit), ("CNPJ/CPF", c_doc), ("EMPRESA", c_name), ("COMPLEMENTARES", c_billing)] if not col]
        if missing:
            raise ValueError("Colunas obrigatórias ausentes: " + ", ".join(missing))
        conn = db()
        conn.execute("BEGIN IMMEDIATE")
        units = {normalize_text(r["name"]): r for r in conn.execute("SELECT * FROM units WHERE active=1").fetchall()}
        exam_columns = {}
        for exam in active_exam_types(conn):
            col = _col_from(h, exam["name"], exam["exam_key"])
            if col:
                exam_columns[exam["exam_key"]] = (exam, col)

        aggregated = {}
        affected_units = set()
        source_rows = 0
        merged_rows = 0
        for ridx in range(2, ws.max_row + 1):
            values = [ws.cell(ridx, col).value for col in range(1, ws.max_column + 1)]
            if all(v is None or str(v).strip() == "" for v in values):
                continue
            source_rows += 1
            try:
                unit = units.get(normalize_text(ws.cell(ridx, c_unit).value))
                if not unit:
                    raise ValueError("unidade não cadastrada ou inativa.")
                doc = _excel_company_document(ws.cell(ridx, c_doc).value)
                if len(doc) not in {11, 14}:
                    raise ValueError("CNPJ/CPF inválido. Use 11 dígitos para CPF ou 14 para CNPJ, com ou sem pontuação.")
                key = (unit["id"], doc)
                existing = conn.execute("SELECT * FROM companies WHERE unit_id=? AND cnpj=?", key).fetchone()
                name = str(ws.cell(ridx, c_name).value or "").strip()
                billing = _billing_mode_excel(ws.cell(ridx, c_billing).value)
                if billing is None:
                    raise ValueError("COMPLEMENTARES inválido. Use SIM ou NÃO.")
                active = _excel_active(ws.cell(ridx, c_active).value if c_active else None, bool(existing["active"]) if existing else True)
                primary_tokens = _import_email_tokens(ws.cell(ridx, c_email).value if c_email else None)
                cc_tokens = _import_email_tokens(ws.cell(ridx, c_cc).value if c_cc else None)
                fixed = parse_money_strict(ws.cell(ridx, c_fixed).value) if c_fixed else None
                prices = {}
                for exam, col in exam_columns.values():
                    price = parse_money_strict(ws.cell(ridx, col).value)
                    if price is not None:
                        prices[exam["exam_key"]] = (exam["name"], price)

                item = aggregated.get(key)
                if item is None:
                    item = {
                        "unit": unit, "doc": doc, "name": name, "existing": existing,
                        "billing": billing, "active": active, "emails": _merge_email_lists(primary_tokens, cc_tokens),
                        "fixed_values": [], "prices": prices,
                        "rows": [ridx],
                    }
                    if fixed is not None:
                        item["fixed_values"].append(fixed)
                    aggregated[key] = item
                else:
                    merged_rows += 1
                    item["rows"].append(ridx)
                    if not item["name"] and name:
                        item["name"] = name
                    if item["billing"] != billing:
                        raise ValueError(f"COMPLEMENTARES diverge das outras linhas do mesmo CNPJ/CPF (linhas {item['rows'][0]} e {ridx}).")
                    if item["active"] != active:
                        raise ValueError(f"ATIVA diverge das outras linhas do mesmo CNPJ/CPF (linhas {item['rows'][0]} e {ridx}).")
                    item["emails"] = _merge_email_lists(item["emails"], primary_tokens, cc_tokens)
                    if fixed is not None:
                        item["fixed_values"].append(fixed)
                    for exam_key, pair in prices.items():
                        if exam_key in item["prices"] and abs(float(item["prices"][exam_key][1]) - float(pair[1])) > 0.0001:
                            raise ValueError(f"preço de {pair[0]} diverge entre linhas do mesmo CNPJ/CPF.")
                        item["prices"][exam_key] = pair
                affected_units.add(unit["name"])
            except ValueError as exc:
                details.append(f"Linha {ridx}: {exc}")

        # Validação final dos grupos consolidados.
        parsed = []
        for item in aggregated.values():
            existing = item["existing"]
            if conn.execute("SELECT 1 FROM complementary_sources WHERE unit_id=? AND document=?", (item["unit"]["id"], item["doc"])).fetchone():
                details.append(f"CNPJ/CPF {item['doc']} está vinculado a uma empresa responsável. Remova o vínculo antes de criar cadastro próprio.")
                continue
            name = item["name"] or (existing["name"] if existing else "")
            if not name:
                details.append(f"Linhas {', '.join(map(str,item['rows']))}: nome da empresa vazio.")
                continue
            emails = item["emails"]
            if emails:
                email = emails[0]
                email_cc = "; ".join(emails[1:])
            elif existing:
                email = existing["email"] or ""
                email_cc = existing["email_cc"] or ""
            else:
                email = ""
                email_cc = ""
            fixed = round(sum(item["fixed_values"]), 2) if item["fixed_values"] else None
            parsed.append((item["unit"], item["doc"], name, email, email_cc, fixed, item["billing"], item["active"], list(item["prices"].values()), existing, item["rows"]))

        if details:
            raise ValueError("A importação contém linhas inválidas. Nenhum cadastro ou preço foi alterado.")
        if not parsed:
            raise ValueError("A planilha não contém empresas para importar.")
        unit_ids = sorted({row[0]["id"] for row in parsed})
        if conn.execute(
            f"SELECT 1 FROM send_jobs j JOIN competencies cp ON cp.id=j.competency_id WHERE cp.unit_id IN ({','.join('?' for _ in unit_ids)}) AND j.status IN ('QUEUED','RUNNING')",
            unit_ids,
        ).fetchone():
            raise ValueError("Há envio em andamento nas unidades da planilha. Aguarde terminar antes de importar cadastros.")

        added = updated = prices_saved = 0
        for unit, doc, name, email, email_cc, fixed, billing, active, prices, row, source_line_numbers in parsed:
            if row:
                conn.execute(
                    "UPDATE companies SET name=?,email=?,email_cc=?,fixed_value=?,bill_complementaries=?,active=?,updated_at=? WHERE id=?",
                    (name, email or row["email"], email_cc if email_cc else row["email_cc"], fixed if fixed is not None else row["fixed_value"], billing, active, now_iso(), row["id"]),
                )
                company_id = row["id"]
                updated += 1
            else:
                company_id = conn.execute(
                    "INSERT INTO companies(unit_id,cnpj,name,email,email_cc,fixed_value,bill_complementaries,active,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                    (unit["id"], doc, name, email, email_cc, fixed if fixed is not None else 0, billing, active, now_iso(), now_iso()),
                ).lastrowid
                added += 1
            if billing and prices:
                for exam_name, price in prices:
                    _upsert_company_price(conn, company_id, exam_name, price)
                    prices_saved += 1
                conn.execute("UPDATE companies SET price_table_id=NULL WHERE id=?", (company_id,))
        conn.commit()
        result_details = []
        if merged_rows:
            result_details.append(f"{merged_rows} linha(s) repetida(s) de CNPJ/CPF foram consolidadas; os valores mensais informados foram somados.")
        result_details.append("E-mails múltiplos encontrados na coluna EMAIL foram separados automaticamente: primeiro endereço como principal e os demais como CC.")
        return {"ok": True, "added": added, "updated": updated, "skipped": 0, "prices_saved": prices_saved,
                "details": result_details, "units": ", ".join(sorted(affected_units)), "exam_columns": len(exam_columns),
                "source_rows": source_rows, "merged_rows": merged_rows, "companies_processed": len(parsed)}
    except Exception as exc:
        if conn is not None:
            conn.rollback()
        return {"ok": False, "error": str(exc), "details": details[:100]}
    finally:
        if conn is not None:
            conn.close()
        if wb is not None:
            wb.close()


def _upsert_company_price(conn, company_id, exam_name, price):
    exam_name=str(exam_name or "").strip()
    if not exam_name:
        return "skipped"
    exam_key=normalize_text(exam_name)
    existing=conn.execute("SELECT id FROM complementary_prices WHERE company_id=? AND exam_key=?",(company_id,exam_key)).fetchone()
    if existing:
        conn.execute("UPDATE complementary_prices SET exam_name=?,unit_price=?,updated_at=? WHERE id=?",(exam_name,price,now_iso(),existing["id"]))
        return "updated"
    conn.execute("INSERT INTO complementary_prices(company_id,exam_name,exam_key,unit_price,created_at,updated_at) VALUES(?,?,?,?,?,?)",(company_id,exam_name,exam_key,price,now_iso(),now_iso()))
    return "added"


def import_prices_workbook(file_storage, unit_id):
    """Importa tabelas e vínculos como um lote validado, mantendo preços vazios."""
    wb = conn = None
    details = []
    try:
        wb = _import_workbook(file_storage)
        conn = db()
        conn.execute("BEGIN IMMEDIATE")
        unit = conn.execute("SELECT * FROM units WHERE id=? AND active=1", (unit_id,)).fetchone()
        if not unit:
            raise ValueError("Selecione uma unidade válida.")
        if conn.execute("SELECT 1 FROM send_jobs j JOIN competencies cp ON cp.id=j.competency_id WHERE cp.unit_id=? AND j.status IN ('QUEUED','RUNNING')", (unit_id,)).fetchone():
            raise ValueError("Há envio em andamento nesta unidade. Aguarde terminar antes de importar preços.")
        ws_tables = _excel_sheet_by_names(wb, "TABELAS_PRECOS", "TABELAS PREÇOS", "TABELAS DE PREÇOS", "TABELAS")
        ws_companies = _excel_sheet_by_names(wb, "EMPRESAS", "MAPEAMENTO", "EMPRESAS_TABELAS")
        if ws_tables is None or ws_companies is None:
            raise ValueError("A planilha precisa ter as abas TABELAS_PRECOS e EMPRESAS.")
        registered = {r["exam_key"]: r for r in active_exam_types(conn)}
        ht = _header_index(ws_tables)
        table_col = _col_from(ht, "TABELA_PRECOS", "TABELA PREÇOS", "TABELA", "NOME DA TABELA")
        if not table_col:
            raise ValueError("A aba TABELAS_PRECOS precisa da coluna TABELA_PRECOS.")
        exam_cols = []
        for col in range(1, ws_tables.max_column + 1):
            if col == table_col:
                continue
            header = str(ws_tables.cell(1, col).value or "").strip()
            if not header:
                continue
            key = normalize_text(header)
            if key not in registered:
                raise ValueError(f"Exame não cadastrado ou inativo em TABELAS_PRECOS: {header}.")
            exam_cols.append((col, registered[key]["name"], key))
        if not exam_cols:
            raise ValueError("A planilha não possui colunas de exames complementares cadastrados.")
        existing_tables = {normalize_text(r["name"]): r for r in conn.execute("SELECT * FROM price_tables WHERE unit_id=?", (unit_id,)).fetchall()}
        price_tables = {}
        for ridx in range(2, ws_tables.max_row + 1):
            values_row = [ws_tables.cell(ridx, col).value for col in range(1, ws_tables.max_column + 1)]
            if all(v is None or str(v).strip() == "" for v in values_row):
                continue
            display = str(ws_tables.cell(ridx, table_col).value or "").strip()
            key = normalize_text(display)
            if not key:
                raise ValueError(f"TABELAS_PRECOS linha {ridx}: nome da tabela vazio.")
            if key in price_tables:
                raise ValueError(f"A tabela '{display}' aparece mais de uma vez na aba TABELAS_PRECOS.")
            values = {}
            for col, exam_name, exam_key in exam_cols:
                try:
                    price = parse_money_strict(ws_tables.cell(ridx, col).value)
                except ValueError as exc:
                    raise ValueError(f"TABELAS_PRECOS linha {ridx}: preço inválido para {exam_name}.") from exc
                if price is not None:
                    values[exam_key] = (exam_name, price)
            existing = existing_tables.get(key)
            if not values and not existing:
                raise ValueError(f"A tabela nova '{display}' não possui preços preenchidos.")
            price_tables[key] = (display, values, existing)
        if not price_tables:
            raise ValueError("A aba TABELAS_PRECOS não contém tabelas para importar.")
        hc = _header_index(ws_companies)
        c_doc = _col_from(hc, "CNPJ/CPF", "CNPJ CPF", "CNPJ", "CPF")
        c_table = _col_from(hc, "TABELA_PRECOS", "TABELA PREÇOS", "TABELA", "TABELA DE PREÇOS")
        if not c_doc or not c_table:
            raise ValueError("A aba EMPRESAS precisa das colunas CNPJ/CPF e TABELA_PRECOS.")
        mappings = []
        seen = set()
        skipped = 0
        for ridx in range(2, ws_companies.max_row + 1):
            raw_doc, raw_table = ws_companies.cell(ridx, c_doc).value, ws_companies.cell(ridx, c_table).value
            if raw_doc in (None, "") and raw_table in (None, ""):
                continue
            doc = _excel_company_document(raw_doc)
            key = normalize_text(raw_table)
            if len(doc) not in {11, 14}:
                raise ValueError(f"EMPRESAS linha {ridx}: CNPJ/CPF inválido.")
            if doc in seen:
                raise ValueError(f"EMPRESAS linha {ridx}: CNPJ/CPF {format_document(doc)} repetido.")
            seen.add(doc)
            company = conn.execute("SELECT * FROM companies WHERE unit_id=? AND cnpj=?", (unit_id, doc)).fetchone()
            if not company:
                raise ValueError(f"EMPRESAS linha {ridx}: {format_document(doc)} não está cadastrado em {unit['name']}.")
            if not company["bill_complementaries"]:
                skipped += 1
                details.append(f"EMPRESAS linha {ridx}: {company['name']} paga no ato; preços mensais mantidos.")
                continue
            if key and key not in price_tables:
                raise ValueError(f"EMPRESAS linha {ridx}: tabela '{raw_table}' não encontrada em TABELAS_PRECOS.")
            if not key:
                details.append(f"EMPRESAS linha {ridx}: tabela vazia; preços atuais de {company['name']} mantidos.")
            mappings.append((company, key))
        before_prices = {r["company_id"]: {} for r in conn.execute("SELECT id company_id FROM companies WHERE unit_id=?", (unit_id,)).fetchall()}
        for r in conn.execute("SELECT p.* FROM complementary_prices p JOIN companies c ON c.id=p.company_id WHERE c.unit_id=?", (unit_id,)).fetchall():
            before_prices[r["company_id"]][r["exam_key"]] = r["unit_price"]
        persistent_ids = {}
        for key, (display, values, existing) in price_tables.items():
            if existing:
                tid = existing["id"]
                conn.execute("UPDATE price_tables SET name=?,active=1,updated_at=? WHERE id=?", (display, now_iso(), tid))
            else:
                tid = conn.execute("INSERT INTO price_tables(unit_id,name,active,created_at,updated_at) VALUES(?,?,1,?,?)", (unit_id, display, now_iso(), now_iso())).lastrowid
            for exam_key, (exam_name, price) in values.items():
                conn.execute("""INSERT INTO price_table_items(table_id,exam_name,exam_key,unit_price,created_at,updated_at)
                    VALUES(?,?,?,?,?,?) ON CONFLICT(table_id,exam_key) DO UPDATE SET exam_name=excluded.exam_name,unit_price=excluded.unit_price,updated_at=excluded.updated_at""",
                    (tid, exam_name, exam_key, price, now_iso(), now_iso()))
            persistent_ids[key] = tid
            sync_price_table_companies(conn, tid)
        for company, key in mappings:
            if key:
                apply_price_table_to_company(conn, company["id"], persistent_ids[key])
        prices_added = prices_updated = 0
        for r in conn.execute("SELECT p.* FROM complementary_prices p JOIN companies c ON c.id=p.company_id WHERE c.unit_id=?", (unit_id,)).fetchall():
            old = before_prices[r["company_id"]]
            if r["exam_key"] not in old:
                prices_added += 1
            elif old[r["exam_key"]] != r["unit_price"]:
                prices_updated += 1
        conn.commit()
        return {"ok": True, "applied": len(mappings), "skipped": skipped, "prices_added": prices_added,
                "prices_updated": prices_updated, "details": details[:100], "unit_name": unit["name"], "tables_count": len(price_tables)}
    except Exception as exc:
        if conn is not None:
            conn.rollback()
        return {"ok": False, "error": str(exc) + " Nenhum preço foi alterado.", "details": details[:100]}
    finally:
        if conn is not None:
            conn.close()
        if wb is not None:
            wb.close()


@app.route("/cadastros")
def registrations():
    conn=db()
    stats=conn.execute("""SELECT
      (SELECT COUNT(*) FROM companies WHERE active=1) companies,
      (SELECT COUNT(*) FROM complementary_exam_types WHERE active=1) exams,
      (SELECT COUNT(*) FROM price_tables WHERE active=1) price_tables,
      (SELECT COUNT(*) FROM units WHERE active=1) units""").fetchone()
    units_rows=conn.execute("SELECT * FROM units WHERE active=1 ORDER BY name").fetchall()
    conn.close()
    return render_template("registrations.html",stats=stats,units=units_rows)


@app.route("/price-tables")
def price_tables():
    unit_id=request.args.get("unit_id",type=int)
    conn=db(); units_rows=conn.execute("SELECT * FROM units WHERE active=1 ORDER BY name").fetchall()
    if not unit_id and units_rows:
        unit_id=units_rows[0]["id"]
    rows=price_tables_for_unit(conn,unit_id,active_only=False) if unit_id else []
    conn.close()
    return render_template("price_tables.html",tables=rows,units=units_rows,selected_unit=unit_id)


@app.route("/price-tables/new",methods=["GET","POST"])
def price_table_new():
    conn = db()
    try:
        units_rows = conn.execute("SELECT * FROM units WHERE active=1 ORDER BY name").fetchall()
        exam_types = active_exam_types(conn)
    finally:
        conn.close()
    selected_unit = request.args.get("unit_id", type=int) or (units_rows[0]["id"] if units_rows else None)
    if request.method == "POST":
        conn = db()
        try:
            conn.execute("BEGIN IMMEDIATE")
            selected_unit = int(request.form.get("unit_id") or 0)
            name = str(request.form.get("name") or "").strip().upper()
            if not name or not conn.execute("SELECT 1 FROM units WHERE id=? AND active=1", (selected_unit,)).fetchone():
                raise ValueError("Informe uma unidade ativa e o nome da tabela.")
            reason = delivery_mutation_reason(conn, unit_id=selected_unit)
            if reason:
                raise ValueError(reason)
            values = []
            for ex in active_exam_types(conn):
                try:
                    price = parse_money_strict(request.form.get(f"price_{ex['id']}"))
                except ValueError as exc:
                    raise ValueError(f"Preço inválido para {ex['name']}.") from exc
                if price is not None:
                    values.append((ex, price))
            tid = conn.execute("INSERT INTO price_tables(unit_id,name,active,created_at,updated_at) VALUES(?,?,1,?,?)", (selected_unit, name, now_iso(), now_iso())).lastrowid
            for ex, price in values:
                conn.execute("INSERT INTO price_table_items(table_id,exam_name,exam_key,unit_price,created_at,updated_at) VALUES(?,?,?,?,?,?)", (tid, ex["name"], ex["exam_key"], price, now_iso(), now_iso()))
            conn.commit()
            audit_event("CRIAR_TABELA_PRECOS", "tabela_precos", tid, f"Tabela de preços {name} criada.", {"unit_id": selected_unit})
            flash("Tabela de preços criada.", "success")
            return redirect(url_for("price_table_edit", table_id=tid))
        except (ValueError, sqlite3.IntegrityError) as exc:
            conn.rollback()
            flash("Já existe uma tabela com esse nome nesta unidade." if isinstance(exc, sqlite3.IntegrityError) else str(exc), "error")
        finally:
            conn.close()
    return render_template("price_table_edit.html", table=None, items={}, exam_types=exam_types, units=units_rows, selected_unit=selected_unit)


@app.route("/price-tables/<int:table_id>/edit",methods=["GET","POST"])
def price_table_edit(table_id):
    conn = db()
    try:
        if request.method == "POST":
            conn.execute("BEGIN IMMEDIATE")
        table = conn.execute("SELECT pt.*,u.name unit_name FROM price_tables pt JOIN units u ON u.id=pt.unit_id WHERE pt.id=?", (table_id,)).fetchone()
        if not table:
            abort(404)
        if request.method == "POST":
            try:
                name = str(request.form.get("name") or "").strip().upper()
                active = int(bool(request.form.get("active")))
                if not name:
                    raise ValueError("Informe o nome da tabela.")
                if int(request.form.get("unit_id") or table["unit_id"]) != table["unit_id"]:
                    raise ValueError("A unidade de uma tabela existente não pode ser alterada.")
                reason = delivery_mutation_reason(conn, unit_id=table["unit_id"])
                if reason:
                    raise ValueError(reason)
                values = []
                for ex in active_exam_types(conn):
                    field = f"price_{ex['id']}"
                    if field not in request.form:
                        continue
                    try:
                        price = parse_money_strict(request.form.get(field))
                    except ValueError as exc:
                        raise ValueError(f"Preço inválido para {ex['name']}.") from exc
                    values.append((ex, price))
                conn.execute("UPDATE price_tables SET name=?,active=?,updated_at=? WHERE id=?", (name, active, now_iso(), table_id))
                for ex, price in values:
                    conn.execute("DELETE FROM price_table_items WHERE table_id=? AND exam_key=?", (table_id, ex["exam_key"]))
                    if price is not None:
                        conn.execute("INSERT INTO price_table_items(table_id,exam_name,exam_key,unit_price,created_at,updated_at) VALUES(?,?,?,?,?,?)", (table_id, ex["name"], ex["exam_key"], price, now_iso(), now_iso()))
                affected = sync_price_table_companies(conn, table_id)
                conn.commit()
                audit_event("ATUALIZAR_TABELA_PRECOS", "tabela_precos", table_id, f"Tabela de preços {name} atualizada.", {"empresas_atualizadas": affected})
                flash(f"Tabela atualizada. {affected} empresa(s) vinculada(s) receberam os novos preços.", "success")
                return redirect(url_for("price_table_edit", table_id=table_id))
            except (ValueError, sqlite3.IntegrityError) as exc:
                conn.rollback()
                flash("Já existe outra tabela com esse nome nesta unidade." if isinstance(exc, sqlite3.IntegrityError) else str(exc), "error")
        table = conn.execute("SELECT pt.*,u.name unit_name FROM price_tables pt JOIN units u ON u.id=pt.unit_id WHERE pt.id=?", (table_id,)).fetchone()
        exam_types = active_exam_types(conn)
        units_rows = conn.execute("SELECT * FROM units WHERE active=1 ORDER BY name").fetchall()
        items = {r["exam_key"]: r for r in conn.execute("SELECT * FROM price_table_items WHERE table_id=?", (table_id,)).fetchall()}
    finally:
        conn.close()
    return render_template("price_table_edit.html", table=table, items=items, exam_types=exam_types, units=units_rows, selected_unit=table["unit_id"])


@app.route("/price-tables/<int:table_id>/delete",methods=["POST"])
def price_table_delete(table_id):
    if not confirm_delete_request():
        flash("Digite EXCLUIR para apagar a tabela de preços.","error"); return redirect(url_for("price_table_edit",table_id=table_id))
    conn=db(); table=conn.execute("SELECT * FROM price_tables WHERE id=?",(table_id,)).fetchone()
    if not table: conn.close(); abort(404)
    create_db_backup(f"antes_excluir_tabela_precos_{table_id}")
    company_ids=[r['id'] for r in conn.execute("SELECT id FROM companies WHERE price_table_id=?",(table_id,)).fetchall()]
    conn.execute("UPDATE companies SET price_table_id=NULL,updated_at=? WHERE price_table_id=?",(now_iso(),table_id))
    conn.execute("DELETE FROM price_table_items WHERE table_id=?",(table_id,)); conn.execute("DELETE FROM price_tables WHERE id=?",(table_id,)); conn.commit(); conn.close()
    audit_event("EXCLUIR_TABELA_PRECOS","tabela_precos",table_id,f"Tabela {table['name']} excluída.",{"empresas_desvinculadas":len(company_ids),"unit_id":table['unit_id']})
    flash(f"Tabela {table['name']} excluída. Os preços já aplicados às empresas foram mantidos.","success"); return redirect(url_for("price_tables",unit_id=table['unit_id']))


@app.route("/exam-types", methods=["GET", "POST"])
def exam_types():
    conn=db()
    if request.method=="POST":
        name=str(request.form.get("name") or "").strip()
        key=normalize_text(name)
        if not name or not key:
            flash("Informe o nome do exame complementar.","error")
        elif key in FORBIDDEN_COMPLEMENTARY_EXAMS:
            flash("Este exame não faz parte do catálogo de complementares da EDGE.","error")
        else:
            existing=conn.execute("SELECT * FROM complementary_exam_types WHERE exam_key=?",(key,)).fetchone()
            if existing:
                if existing["active"]:
                    flash("Este tipo de exame já está cadastrado.","warning")
                else:
                    conn.execute("UPDATE complementary_exam_types SET name=?,active=1,updated_at=? WHERE id=?",(name,now_iso(),existing["id"]))
                    conn.commit(); flash("Tipo de exame reativado.","success")
            else:
                cur=conn.execute("INSERT INTO complementary_exam_types(name,exam_key,active,created_at,updated_at) VALUES(?,?,1,?,?)",(name,key,now_iso(),now_iso()))
                conn.commit(); audit_event("CRIAR_TIPO_EXAME","tipo_exame",cur.lastrowid,f"Tipo de exame {name} cadastrado."); flash("Tipo de exame cadastrado.","success")
    forbidden = tuple(FORBIDDEN_COMPLEMENTARY_EXAMS)
    rows=conn.execute("""SELECT e.*,
      (SELECT COUNT(*) FROM complementary_prices p WHERE p.exam_key=e.exam_key) price_count
      FROM complementary_exam_types e
      WHERE e.exam_key NOT IN (?,?)
      ORDER BY e.active DESC,e.name""", forbidden).fetchall()
    conn.close()
    return render_template("exam_types.html",exam_types=rows)


@app.route("/exam-types/<int:exam_id>/toggle", methods=["POST"])
def exam_type_toggle(exam_id):
    conn=db(); row=conn.execute("SELECT * FROM complementary_exam_types WHERE id=?",(exam_id,)).fetchone()
    if not row:
        conn.close(); abort(404)
    if row["exam_key"] in FORBIDDEN_COMPLEMENTARY_EXAMS:
        conn.close(); flash("Este exame não faz parte do catálogo de complementares da EDGE.","error"); return redirect(url_for("exam_types"))
    new_active=0 if row["active"] else 1
    conn.execute("UPDATE complementary_exam_types SET active=?,updated_at=? WHERE id=?",(new_active,now_iso(),exam_id))
    conn.commit(); conn.close(); audit_event("ATIVAR_TIPO_EXAME" if new_active else "DESATIVAR_TIPO_EXAME","tipo_exame",exam_id,f"Tipo de exame {row['name']} {'ativado' if new_active else 'desativado'}.")
    flash("Tipo de exame ativado." if new_active else "Tipo de exame desativado. Ele não aparecerá em novas planilhas de preços.","success" if new_active else "warning")
    return redirect(url_for("exam_types"))


@app.route("/exam-types/<int:exam_id>/delete", methods=["POST"])
def exam_type_delete(exam_id):
    if not confirm_delete_request():
        flash("Digite EXCLUIR para apagar definitivamente o tipo de exame.","error")
        return redirect(url_for("exam_types"))
    conn=db(); row=conn.execute("SELECT * FROM complementary_exam_types WHERE id=?",(exam_id,)).fetchone()
    if not row:
        conn.close(); abort(404)
    price_count=conn.execute("SELECT COUNT(*) n FROM complementary_prices WHERE exam_key=?",(row["exam_key"],)).fetchone()["n"]
    table_item_count=conn.execute("SELECT COUNT(*) n FROM price_table_items WHERE exam_key=?",(row["exam_key"],)).fetchone()["n"]
    try:
        create_db_backup(f"antes_excluir_exame_{exam_id}")
        conn.execute("BEGIN IMMEDIATE")
        conn.execute("DELETE FROM complementary_prices WHERE exam_key=?",(row["exam_key"],))
        conn.execute("DELETE FROM price_table_items WHERE exam_key=?",(row["exam_key"],))
        conn.execute("DELETE FROM complementary_exam_types WHERE id=?",(exam_id,))
        conn.commit()
    except Exception as exc:
        conn.rollback(); conn.close(); flash(f"Não foi possível excluir o exame: {exc}","error"); return redirect(url_for("exam_types"))
    conn.close(); audit_event("EXCLUIR_TIPO_EXAME","tipo_exame",exam_id,f"Tipo de exame {row['name']} excluído.",{"precos_empresas":price_count,"itens_tabela":table_item_count})
    flash(f"Tipo de exame '{row['name']}' excluído definitivamente. {price_count} preço(s) de empresas e {table_item_count} item(ns) de tabela foram removidos. O histórico já apurado permanece.","success")
    return redirect(url_for("exam_types"))


@app.route("/companies")
def companies():
    unit_id = request.args.get("unit_id", type=int)
    q = str(request.args.get("q") or "").strip()
    selected_status = request.args.get("status", "active")
    if selected_status not in {"active", "archived", "all"}:
        selected_status = "active"
    conn = db()
    try:
        params = []; where = []
        if unit_id:
            where.append("c.unit_id=?"); params.append(unit_id)
        if selected_status != "all":
            where.append("c.active=?"); params.append(1 if selected_status == "active" else 0)
        if q:
            terms = ["c.name LIKE ?", "c.cnpj LIKE ?", "c.email LIKE ?"]
            like = f"%{q}%"; params += [like, like, like]
            document_query = digits(q)
            if len(document_query) >= 3:
                terms.append("c.cnpj LIKE ?"); params.append(f"%{document_query}%")
            where.append("(" + " OR ".join(terms) + ")")
        sql = """SELECT c.*,COALESCE(u.name,'Sem unidade válida') unit_name,pt.name price_table_name,
                 (SELECT COUNT(*) FROM complementary_prices p WHERE p.company_id=c.id) price_count
                 FROM companies c LEFT JOIN units u ON u.id=c.unit_id
                 LEFT JOIN price_tables pt ON pt.id=c.price_table_id"""
        if where: sql += " WHERE " + " AND ".join(where)
        rows = conn.execute(sql + " ORDER BY u.name,c.name", params).fetchall()
        units_rows = conn.execute("SELECT * FROM units ORDER BY name").fetchall()
        selected_unit_name = next((u["name"] for u in units_rows if u["id"] == unit_id), None)
        counts = conn.execute("SELECT COUNT(*) total,SUM(CASE WHEN active=1 THEN 1 ELSE 0 END) active FROM companies" + (" WHERE unit_id=?" if unit_id else ""), (unit_id,) if unit_id else ()).fetchone()
        total_all_units = conn.execute("SELECT COUNT(*) FROM companies").fetchone()[0]
    finally:
        conn.close()
    return render_template("companies.html", companies=rows, units=units_rows, selected_unit=unit_id, selected_unit_name=selected_unit_name,
                           selected_status=selected_status, q=q, scope_total=counts["total"], scope_active=counts["active"] or 0,
                           scope_archived=counts["total"]-(counts["active"] or 0), total_all_units=total_all_units)


@app.route("/companies/bulk-delete/review", methods=["POST"])
def companies_bulk_delete_review():
    import company_deletions
    unit_id = request.form.get("scope_unit_id", type=int)
    mode = request.form.get("mode", "selected")
    conn = db()
    try:
        if request.form.get("scope_unit_id") and (unit_id is None or unit_id <= 0):
            raise ValueError("Escolha uma unidade válida ou todas as unidades.")
        conn.execute("BEGIN")
        ids = company_deletions.select_ids(conn, mode, request.form.getlist("company_ids"), unit_id)
        plan = company_deletions.build_plan(sys.modules[__name__], conn, ids, unit_id)
        token = company_deletions.encode_plan(sys.modules[__name__], plan, mode)
    except (ValueError, sqlite3.Error) as exc:
        flash(f"Não foi possível preparar a revisão: {exc}", "error")
        return redirect(url_for("companies", unit_id=unit_id))
    finally:
        conn.rollback(); conn.close()
    return render_template("companies_delete_review.html", plan=plan, deletion_token=token, mode=mode)


@app.route("/companies/bulk-delete", methods=["POST"])
def companies_bulk_delete():
    import company_deletions
    unit_id = None
    try:
        reviewed = company_deletions.decode_plan(sys.modules[__name__], request.form.get("deletion_token", ""))
        unit_id = reviewed.get("unit_id")
        if normalize_text(request.form.get("confirm_text")) != f"APAGAR {len(reviewed['ids'])}":
            raise ValueError(f"Digite APAGAR {len(reviewed['ids'])} para confirmar a quantidade revisada.")
        result = company_deletions.execute(sys.modules[__name__], reviewed["ids"], unit_id, reviewed["fingerprint"])
        flash(f"{result['delete_count']} empresa(s) apagada(s) sem histórico; {result['archive_count']} arquivada(s) com histórico preservado. Backup completo criado. {len(result['groups'])} grupo(s) atualizado(s) ou desfeito(s).", "success")
        for warning in result["warnings"]:
            flash(warning, "warning")
    except Exception as exc:
        flash(f"Nenhuma exclusão confirmada: {exc}", "error")
    return redirect(url_for("companies", unit_id=unit_id))



@app.route("/companies/import", methods=["GET", "POST"])
def companies_import():
    result = None
    if request.method == "POST":
        f = request.files.get("file")
        if not f or not f.filename:
            flash("Selecione uma planilha Excel.", "error")
        elif Path(f.filename).suffix.lower() not in {".xlsx", ".xlsm"}:
            flash("Envie um arquivo .xlsx ou .xlsm.", "error")
        else:
            result = import_companies_workbook(f)
            if result.get("ok"):
                audit_event("IMPORTAR_EMPRESAS","cadastro",None,f"Cadastro completo: {result['added']} nova(s), {result['updated']} atualizada(s), {result['prices_saved']} preço(s).",result)
                flash(f"Cadastro completo importado: {result['added']} nova(s), {result['updated']} atualizada(s) e {result['prices_saved']} preço(s) gravado(s).", "success")
            else:
                flash(result.get("error") or "Não foi possível importar.", "error")
    return render_template("companies_import.html", result=result)


@app.route("/companies/import/model")
def companies_import_model():
    conn = db()
    units = conn.execute("SELECT * FROM units WHERE active=1 ORDER BY name").fetchall()
    exams = active_exam_types(conn)
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "EMPRESAS"
    headers = ["UNIDADE", "CNPJ/CPF", "EMPRESA", "EMAIL", "EMAIL_CC", "VALOR_FIXO_MENSAL", "ATIVA", "COMPLEMENTARES"] + [e["name"] for e in exams]
    ws.append(headers)
    unit_names = [u["name"] for u in units]
    if unit_names and len(",".join(unit_names)) < 240:
        dv_unit = DataValidation(type="list", formula1='"' + ",".join(unit_names) + '"', allow_blank=False)
        ws.add_data_validation(dv_unit); dv_unit.add("A2:A2000")
    dv_yesno = DataValidation(type="list", formula1='"SIM,NÃO"', allow_blank=False)
    ws.add_data_validation(dv_yesno); dv_yesno.add("G2:G2000"); dv_yesno.add("H2:H2000")
    for c in ws[1]:
        c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="D9EAF7" if c.column <= 8 else "FFF2CC"); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    widths = {1:16, 2:22, 3:42, 4:32, 5:32, 6:20, 7:12, 8:20}
    for idx in range(1, ws.max_column + 1): ws.column_dimensions[get_column_letter(idx)].width = widths.get(idx, 24)

    we = wb.create_sheet("EXEMPLO")
    we.append(headers)
    sample_prices = []
    for e in exams:
        key = normalize_text(e["name"])
        sample_prices.append(60 if key == "AUDIOMETRIA" else 80 if key in {"ESPIROMETRIA", "ACUIDADE VISUAL"} else 30 if key == "ANAMNESE PSICOSSOCIAL OCUPACIONAL" else 120 if key == "LAUDO PCD" else None)
    we.append([unit_names[0] if unit_names else "BELÉM", "12.345.678/0001-90", "EMPRESA EXEMPLO LTDA", "financeiro@empresa.com.br", "rh@empresa.com.br", 500, "SIM", "SIM"] + sample_prices)
    we.append([unit_names[0] if unit_names else "BELÉM", "123.456.789-01", "CLIENTE PESSOA FÍSICA", "cliente@email.com", "", 350, "SIM", "NÃO"] + [None for _ in exams])
    for c in we[1]:
        c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="D9EAF7" if c.column <= 8 else "FFF2CC"); c.alignment = Alignment(horizontal="center", wrap_text=True)
    for idx in range(1, we.max_column + 1): we.column_dimensions[get_column_letter(idx)].width = widths.get(idx, 24)
    we.freeze_panes = "A2"

    wi = wb.create_sheet("INSTRUÇÕES")
    rows = [
        ["CAMPO", "COMO PREENCHER"],
        ["UNIDADE", "Informe uma unidade já cadastrada no sistema, por exemplo BELÉM ou MACAPÁ."],
        ["CNPJ/CPF", "Pessoa jurídica: CNPJ. Pessoa física: CPF. Pode ser informado com ou sem pontuação."],
        ["EMPRESA", "Nome da empresa ou da pessoa física usada na cobrança."],
        ["EMAIL", "E-mail principal que receberá a cobrança."],
        ["EMAIL_CC", "Opcional. Para vários e-mails, separe por ponto e vírgula."],
        ["VALOR_FIXO_MENSAL", "Valor da mensalidade fixa."],
        ["ATIVA", "SIM para ativa; NÃO para inativa."],
        ["COMPLEMENTARES", "SIM = cobrança no fechamento mensal. NÃO = complementares pagos no ato."],
        ["PREÇOS", "Preencha os valores unitários dos complementares cobrados da empresa. Célula vazia = sem preço/sem contratação."],
        ["COMPLEMENTARES = NÃO", "Deixe as colunas de preços em branco."],
        ["ATUALIZAÇÃO", "Se CNPJ/CPF já existir na mesma unidade, a empresa e seus preços serão atualizados, sem duplicar."],
    ]
    for row in rows: wi.append(row)
    for c in wi[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E78")
    wi.column_dimensions["A"].width = 34; wi.column_dimensions["B"].width = 100
    for row in wi.iter_rows():
        for c in row: c.alignment = Alignment(vertical="top", wrap_text=True)
    wi.freeze_panes = "A2"

    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return send_file(bio, as_attachment=True, download_name="MODELO_CADASTRO_COMPLETO_EMPRESAS_COBRANCAS.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/companies/prices/import", methods=["GET", "POST"])
def prices_import():
    flash("Empresas e preços são importados juntos na planilha de cadastro completo.", "info")
    return redirect(url_for("companies_import"))


@app.route("/companies/prices/model")
def prices_import_model():
    return redirect(url_for("companies_import_model"))


def _company_form_values(conn, form, company=None):
    try:
        unit_id = int(form.get("unit_id") or (company["unit_id"] if company else 0))
        table_id = int(form.get("price_table_id") or 0) or None
    except (TypeError, ValueError) as exc:
        raise ValueError("Selecione uma unidade e uma tabela de preços válidas.") from exc
    if not conn.execute("SELECT 1 FROM units WHERE id=? AND active=1", (unit_id,)).fetchone():
        raise ValueError("Selecione uma unidade ativa.")
    for checked_unit in {unit_id, company["unit_id"] if company else unit_id}:
        reason = delivery_mutation_reason(conn, unit_id=checked_unit)
        if reason:
            raise ValueError(reason)
    if company and company["unit_id"] != unit_id and conn.execute("SELECT 1 FROM billing_group_members WHERE company_id=?", (company["id"],)).fetchone():
        raise ValueError("Desfaça o vínculo com o grupo de cobrança antes de mover a empresa para outra unidade.")
    document = digits(form.get("cnpj"))
    if conn.execute("SELECT 1 FROM complementary_sources WHERE unit_id=? AND document=?", (unit_id, document)).fetchone():
        raise ValueError("Este CNPJ/CPF já está vinculado a uma empresa responsável. Remova o vínculo antes de criar cadastro próprio.")
    name = str(form.get("name") or "").strip()
    email = str(form.get("email") or "").strip().lower()
    email_cc = str(form.get("email_cc") or "").strip().lower()
    if len(document) not in {11, 14} or not name:
        raise ValueError("Informe CNPJ/CPF válido e nome da empresa.")
    if email and not valid_email(email):
        raise ValueError("E-mail principal inválido.")
    try:
        fixed = parse_money_strict(form.get("fixed_value"))
    except ValueError as exc:
        raise ValueError("Valor fixo mensal inválido. Informe um valor não negativo.") from exc
    if fixed is None:
        fixed = company["fixed_value"] if company else 0
    billing = int(bool(form.get("bill_complementaries")))
    if billing and table_id:
        if not conn.execute("SELECT 1 FROM price_tables WHERE id=? AND unit_id=? AND active=1", (table_id, unit_id)).fetchone():
            raise ValueError("A tabela de preços deve estar ativa e pertencer à unidade da empresa.")
    else:
        table_id = None
    return unit_id, document, name, email, email_cc, fixed, billing, table_id


@app.route("/companies/new", methods=["GET", "POST"])
def company_new():
    conn = db()
    try:
        if request.method == "POST":
            try:
                conn.execute("BEGIN IMMEDIATE")
                unit_id, document, name, email, email_cc, fixed, billing, table_id = _company_form_values(conn, request.form)
                company_id = conn.execute("INSERT INTO companies(unit_id,cnpj,name,email,email_cc,fixed_value,bill_complementaries,active,created_at,updated_at,price_table_id) VALUES(?,?,?,?,?,?,?,1,?,?,?)", (unit_id, document, name, email, email_cc, fixed, billing, now_iso(), now_iso(), table_id)).lastrowid
                if billing and table_id:
                    apply_price_table_to_company(conn, company_id, table_id)
                elif billing:
                    save_price_form(conn, company_id, request.form)
                billing_policy.save_company(sys.modules[__name__], conn, company_id, request.form)
                conn.commit()
                audit_event("CADASTRAR_EMPRESA", "empresa", company_id, f"Empresa {name} cadastrada.", {"documento": document, "fixed_value": fixed, "bill_complementaries": billing}, company_id=company_id)
                flash("Empresa cadastrada.", "success")
                return redirect(url_for("company_edit", company_id=company_id))
            except (ValueError, sqlite3.IntegrityError) as exc:
                conn.rollback()
                flash("Este CNPJ/CPF já está cadastrado nesta unidade." if isinstance(exc, sqlite3.IntegrityError) else str(exc), "error")
        units_rows = conn.execute("SELECT * FROM units WHERE active=1 ORDER BY name").fetchall()
        exam_types = active_exam_types(conn)
        price_tables_all = conn.execute("SELECT * FROM price_tables WHERE active=1 ORDER BY unit_id,name").fetchall()
    finally:
        conn.close()
    return render_template("company_edit.html", company=None, prices=[], price_values={}, price_records={}, exam_types=exam_types, units=units_rows, price_tables=price_tables_all, sources_text=request.form.get("complementary_sources", ""))


def save_price_form(conn, company_id, form):
    names, values = form.getlist("exam_name"), form.getlist("exam_price")
    if len(names) != len(values):
        raise ValueError("A lista de exames e preços está incompleta. Recarregue o formulário.")
    registered = {r["exam_key"]: r for r in active_exam_types(conn)}
    parsed = {}
    for name, value in zip(names, values):
        key = normalize_text(name)
        if key not in registered:
            raise ValueError(f"Exame não cadastrado ou inativo: {name}.")
        if key in parsed:
            raise ValueError(f"Exame repetido no formulário: {name}.")
        try:
            price = parse_money_strict(value)
        except ValueError as exc:
            raise ValueError(f"Preço inválido para {registered[key]['name']}.") from exc
        parsed[key] = (registered[key]["name"], price)
    # Somente os campos enviados são alterados; preços inativos são preservados.
    for key, (name, price) in parsed.items():
        conn.execute("DELETE FROM complementary_prices WHERE company_id=? AND exam_key=?", (company_id, key))
        if price is not None:
            _upsert_company_price(conn, company_id, name, price)


@app.route("/companies/<int:company_id>/edit", methods=["GET", "POST"])
def company_edit(company_id):
    conn = db()
    try:
        if request.method == "POST":
            conn.execute("BEGIN IMMEDIATE")
        company = conn.execute("SELECT c.*,u.name unit_name FROM companies c JOIN units u ON u.id=c.unit_id WHERE c.id=?", (company_id,)).fetchone()
        if not company:
            abort(404)
        if request.method == "POST":
            try:
                unit_id, document, name, email, email_cc, fixed, billing, table_id = _company_form_values(conn, request.form, company)
                active = int(bool(request.form.get("active")))
                conn.execute("UPDATE companies SET unit_id=?,cnpj=?,name=?,email=?,email_cc=?,fixed_value=?,bill_complementaries=?,active=?,price_table_id=?,updated_at=? WHERE id=?", (unit_id, document, name, email, email_cc, fixed, billing, active, table_id, now_iso(), company_id))
                if billing and table_id:
                    apply_price_table_to_company(conn, company_id, table_id)
                elif billing:
                    save_price_form(conn, company_id, request.form)
                else:
                    conn.execute("DELETE FROM complementary_prices WHERE company_id=?", (company_id,))
                billing_policy.save_company(sys.modules[__name__], conn, company_id, request.form)
                conn.commit()
                audit_event("ATUALIZAR_EMPRESA", "empresa", company_id, f"Cadastro e preços de {name} atualizados.", {"documento": document, "fixed_value": fixed, "bill_complementaries": billing}, company_id=company_id)
                flash("Empresa atualizada. Se mudou os CNPJs vinculados, reprocesse o controle completo do mês para atualizar as cobranças ainda não enviadas.", "success")
                return redirect(url_for("company_edit", company_id=company_id))
            except (ValueError, sqlite3.IntegrityError) as exc:
                conn.rollback()
                flash("Já existe empresa com este CNPJ/CPF nesta unidade." if isinstance(exc, sqlite3.IntegrityError) else str(exc), "error")
        company = conn.execute("SELECT c.*,u.name unit_name FROM companies c JOIN units u ON u.id=c.unit_id WHERE c.id=?", (company_id,)).fetchone()
        units_rows = conn.execute("SELECT * FROM units WHERE active=1 ORDER BY name").fetchall()
        prices = conn.execute("SELECT * FROM complementary_prices WHERE company_id=? ORDER BY exam_name", (company_id,)).fetchall()
        exam_types = active_exam_types(conn)
        price_tables_all = conn.execute("SELECT * FROM price_tables WHERE active=1 ORDER BY unit_id,name").fetchall()
        price_values = {r["exam_key"]: r["unit_price"] for r in prices}
        price_records = {r["exam_key"]: dict(r) for r in prices}
        sources_text = billing_policy.sources_text(conn, company_id)
    finally:
        conn.close()
    return render_template("company_edit.html", company=company, prices=prices, price_values=price_values, price_records=price_records, exam_types=exam_types, units=units_rows, price_tables=price_tables_all, sources_text=request.form.get("complementary_sources", sources_text))


@app.route("/companies/<int:company_id>/delete", methods=["POST"])
def company_delete(company_id):
    import company_deletions
    if not confirm_delete_request():
        flash("Digite EXCLUIR para retirar a empresa do cadastro ativo.", "error")
        return redirect(url_for("company_edit", company_id=company_id))
    conn = db()
    try:
        company = conn.execute("SELECT * FROM companies WHERE id=?", (company_id,)).fetchone()
        unit_id = company["unit_id"] if company and conn.execute("SELECT 1 FROM units WHERE id=?", (company["unit_id"],)).fetchone() else None
    finally:
        conn.close()
    if not company:
        abort(404)
    try:
        result = company_deletions.execute(sys.modules[__name__], [company_id], unit_id)
        if result["archive_count"]:
            flash(f"Empresa {company['name']} arquivada e retirada da lista ativa. Cobranças, pagamentos, anexos e recibos de envio foram preservados.", "success")
        else:
            flash(f"Empresa {company['name']} apagada, junto dos rascunhos sem emissão ou pagamento. Backup completo criado.", "success")
        for warning in result["warnings"]:
            flash(warning, "warning")
    except Exception as exc:
        flash(f"Não foi possível retirar a empresa: {exc}", "error")
        return redirect(url_for("company_edit", company_id=company_id))
    return redirect(url_for("companies", unit_id=unit_id))


@app.route("/companies/<int:company_id>/prices/<int:price_id>/delete", methods=["POST"])
def company_price_delete(company_id, price_id):
    conn=db(); row=conn.execute("SELECT * FROM complementary_prices WHERE id=? AND company_id=?",(price_id,company_id)).fetchone()
    if not row:
        conn.close(); abort(404)
    create_db_backup(f"antes_excluir_preco_{price_id}")
    conn.execute("DELETE FROM complementary_prices WHERE id=?",(price_id,)); conn.commit(); conn.close()
    audit_event("EXCLUIR_PRECO","empresa",company_id,f"Preço de {row['exam_name']} removido.",company_id=company_id)
    flash(f"Preço de {row['exam_name']} removido da empresa.","success")
    return redirect(url_for("company_edit",company_id=company_id))


@app.route("/competencies/new", methods=["GET", "POST"])
def competency_new():
    conn = db()
    unit_id = request.args.get("unit_id", type=int)
    try:
        if request.method == "POST":
            try:
                try:
                    unit_id = int(request.form.get("unit_id") or 0)
                    month = int(request.form.get("month") or 0)
                    year = int(request.form.get("year") or 0)
                except (TypeError, ValueError) as exc:
                    raise ValueError("Informe unidade, mês e ano válidos.") from exc
                if month not in range(1, 13) or not 2020 <= year <= 9999:
                    raise ValueError("Competência inválida. Informe mês de 1 a 12 e ano entre 2020 e 9999.")
                conn.execute("BEGIN IMMEDIATE")
                if not conn.execute("SELECT 1 FROM units WHERE id=? AND active=1", (unit_id,)).fetchone():
                    raise ValueError("Selecione uma unidade ativa.")
                existing = conn.execute("SELECT id FROM competencies WHERE unit_id=? AND month=? AND year=?", (unit_id, month, year)).fetchone()
                if existing:
                    conn.rollback()
                    flash("Essa competência já existe.", "warning")
                    return redirect(url_for("competency_detail", competency_id=existing["id"]))
                cid = conn.execute("INSERT INTO competencies(unit_id,month,year,created_at,updated_at) VALUES(?,?,?,?,?)", (unit_id, month, year, now_iso(), now_iso())).lastrowid
                sync_competency_companies(conn, cid)
                conn.commit()
                flash("Competência criada. As empresas ativas já foram carregadas automaticamente.", "success")
                return redirect(url_for("competency_apuration_page", competency_id=cid))
            except (TypeError, ValueError, sqlite3.IntegrityError) as exc:
                conn.rollback()
                flash("Informe unidade, mês e ano válidos." if isinstance(exc, (TypeError, sqlite3.IntegrityError)) else str(exc), "error")
        units_rows = conn.execute("SELECT * FROM units WHERE active=1 ORDER BY name").fetchall()
    finally:
        conn.close()
    return render_template("competency_new.html", units=units_rows, selected_unit=unit_id)


@app.route("/competencies/<int:competency_id>")
def competency_detail(competency_id):
    selected_filter=str(request.args.get("filter") or "all")
    q=str(request.args.get("q") or "")
    snap=competency_operational_snapshot(competency_id,selected_filter,q)
    if not snap: abort(404)
    return render_template("competency_detail.html",**snap)


@app.route("/competencies/<int:competency_id>/apurar")
def competency_apuration_page(competency_id):
    snap=competency_operational_snapshot(competency_id)
    if not snap: abort(404)
    unpriced=[x for x in snap["all_items"] if x["comp_state"]["code"]=="UNPRICED"]
    conn=db()
    last_run=conn.execute("SELECT * FROM apuration_runs WHERE competency_id=? ORDER BY id DESC LIMIT 1",(competency_id,)).fetchone()
    duplicate_items=conn.execute("""SELECT e.*,c.name company_name,c.cnpj FROM exam_items e JOIN companies c ON c.id=e.company_id
                                    WHERE e.competency_id=? AND e.status='JA_COBRADO' ORDER BY c.name,e.employee,e.exam_name""",(competency_id,)).fetchall()
    conn.close()
    last_stats={}
    if last_run:
        try: last_stats=json.loads(last_run["stats_json"] or "{}")
        except Exception: last_stats={}
    return render_template("competency_apuration_page.html",**snap,unpriced_companies=unpriced,last_run=last_run,last_stats=last_stats,duplicate_items=duplicate_items)


@app.route("/competencies/<int:competency_id>/documents")
def competency_documents_page(competency_id):
    snap=competency_operational_snapshot(competency_id)
    if not snap: abort(404)
    conn=db()
    fixed_docs=conn.execute("SELECT COUNT(*) FROM documents WHERE competency_id=? AND billing_type='FIXED'",(competency_id,)).fetchone()[0]
    comp_docs=conn.execute("SELECT COUNT(*) FROM documents WHERE competency_id=? AND billing_type='COMPLEMENTARY'",(competency_id,)).fetchone()[0]
    conn.close()
    fixed_missing=sum(1 for x in snap["all_items"] if x["fixed_state"]["code"]=="NO_DOC")
    comp_missing=sum(1 for x in snap["all_items"] if x["comp_state"]["code"]=="NO_DOC")
    conn=db(); last_doc_runs=conn.execute("SELECT * FROM document_import_runs WHERE competency_id=? ORDER BY id DESC LIMIT 6",(competency_id,)).fetchall(); conn.close()
    return render_template("competency_documents_page.html",**snap,fixed_docs_total=fixed_docs,comp_docs_total=comp_docs,fixed_missing=fixed_missing,comp_missing=comp_missing,last_doc_runs=last_doc_runs)


def combined_review_data(conn, competency_id):
    ready, blocked = [], []
    uncertain = _send_queue.uncertain_members(conn, competency_id, 'COMBINED')
    for envelope in combined_billing.envelopes(conn, competency_id):
        payload = combined_billing.build_payload(sys.modules[__name__], conn, competency_id, envelope['owner'], envelope=envelope)
        if not payload: continue
        problem = combined_billing.reason(sys.modules[__name__], conn, competency_id, envelope)
        if uncertain.intersection(envelope['ids']):
            problem = 'Há envio com resultado incerto. Confira o Gmail e resolva a fila anterior.'
        if problem:
            blocked.append(dict(company_id=envelope['owner'], label=payload['delivery_label'], reason=problem))
        else:
            payload = complementary_report.enrich(sys.modules[__name__], conn, competency_id, payload)
            ready.append(payload)
    return dict(combined_items=ready, combined_blocked=blocked,
        combined_fixed_total=sum(sum(parts.get('FIXED',0) for parts in item['financial_components'].values()) for item in ready),
        combined_comp_total=sum(sum(parts.get('COMPLEMENTARY',0) for parts in item['financial_components'].values()) for item in ready))


def resend_review_data(conn, competency_id):
    """Monta os pacotes já enviados e ainda não pagos que podem ser reenviados.

    O reenvio usa exatamente a mesma regra de destinatário/grupo da cobrança original.
    `force=True` significa *somente já enviado*, evitando incluir empresas novas no lote.
    """
    result = {}
    for billing, prefix, label in (
        ("FIXED", "fixed", "Mensalidade"),
        ("COMPLEMENTARY", "complementary", "Complementares"),
    ):
        sql = f"""SELECT cc.company_id FROM competency_companies cc
            JOIN companies c ON c.id=cc.company_id
            WHERE cc.competency_id=? AND c.active=1
              AND cc.{prefix}_sent_at IS NOT NULL AND cc.{prefix}_paid=0
              {"AND c.bill_complementaries=1" if billing == "COMPLEMENTARY" else ""}
            ORDER BY c.name,c.id"""
        candidates = conn.execute(sql, (competency_id,)).fetchall()
        ready, blocked, seen = [], [], set()
        uncertain = _send_queue.uncertain_members(conn, competency_id, billing)
        for candidate in candidates:
            representative = billing_delivery_representative(conn, competency_id, billing, candidate["company_id"], force=True)
            if representative in seen:
                continue
            seen.add(representative)
            member_ids = billing_delivery_ids(conn, competency_id, billing, representative, force=True)
            if not member_ids:
                continue
            reason = billing_send_reason(conn, competency_id, billing, representative, force=True)
            if uncertain.intersection(member_ids):
                reason = "Há envio com resultado incerto. Confira o Gmail antes de reenviar."
            marks = ",".join("?" for _ in member_ids)
            members = conn.execute(
                f"SELECT id,name,cnpj,email FROM companies WHERE id IN ({marks}) ORDER BY name", tuple(member_ids)
            ).fetchall()
            amount = conn.execute(
                f"SELECT COALESCE(SUM({prefix}_amount),0) total FROM competency_companies WHERE competency_id=? AND company_id IN ({marks})",
                (competency_id, *member_ids),
            ).fetchone()["total"]
            owner = conn.execute("SELECT id,name,email FROM companies WHERE id=?", (representative,)).fetchone()
            item = dict(
                company_id=representative,
                owner_name=owner["name"] if owner else "Empresa",
                email=owner["email"] if owner else "",
                member_ids=member_ids,
                member_names=[m["name"] for m in members],
                member_count=len(member_ids),
                amount=float(amount or 0),
                billing=billing,
                label=label,
                reason=reason,
            )
            (blocked if reason else ready).append(item)
        result[billing] = {"ready": ready, "blocked": blocked}
    return result


@app.route("/competencies/<int:competency_id>/review")
def competency_review(competency_id):
    snap=competency_operational_snapshot(competency_id)
    if not snap: abort(404)
    conn=db()
    try:
        combined=combined_review_data(conn,competency_id)
        resends=resend_review_data(conn,competency_id)
        errors=conn.execute("SELECT * FROM send_jobs WHERE competency_id=? AND error_groups>0 AND status IN ('COMPLETED_WITH_ERRORS','FAILED') ORDER BY created_at DESC LIMIT 5",(competency_id,)).fetchall()
    finally: conn.close()
    return render_template('competency_review.html',**snap,**combined,resends=resends,last_error_jobs=errors)


def billing_review_item(payload,billing):
    item=dict(payload["row"])
    members=payload.get("members") or []
    batch=bool(payload.get("batched") or payload.get("recipient_batched"))
    item.update(billing_group_name=payload.get("group_name"),billing_members=members,
                recipient_batched=batch,delivery_label=f"{len(members)} empresas no mesmo e-mail" if batch else item["name"],
                company_names=" · ".join(member["name"] for member in members))
    item["fixed_amount" if billing=="FIXED" else "complementary_amount"]=payload["amount"]
    item["fixed_docs" if billing=="FIXED" else "comp_docs"]=len(payload["attachments"])
    return item


@app.route("/competencies/<int:competency_id>/payments-view")
def competency_payments_page(competency_id):
    snap=competency_operational_snapshot(competency_id)
    if not snap: abort(404)
    payment_items=[x for x in snap["all_items"] if x["fixed_sent_at"] or x["complementary_sent_at"]]
    conn=db()
    last_payment_run=conn.execute("SELECT * FROM payment_import_runs WHERE competency_id=? ORDER BY id DESC LIMIT 1",(competency_id,)).fetchone()
    conn.close()
    return render_template("competency_payments_page.html",**snap,payment_items=payment_items,last_payment_run=last_payment_run)


@app.route("/competencies/<int:competency_id>/pending")
def competency_pending(competency_id):
    # Compatibilidade com links antigos: a Central de Pendências virou a tela principal da competência na V2.
    return redirect(url_for("competency_detail", competency_id=competency_id, filter=request.args.get("filter") or "all", q=request.args.get("q") or ""))


@app.route("/competencies/<int:competency_id>/close",methods=["POST"])
def competency_close(competency_id):
    comp=get_competency(competency_id)
    if not comp: abort(404)
    conn=db(); running=conn.execute("SELECT 1 FROM send_jobs WHERE competency_id=? AND status IN ('QUEUED','RUNNING')",(competency_id,)).fetchone(); conn.close()
    if running:
        flash("Aguarde o envio em andamento terminar antes de fechar a competência.","error"); return redirect(url_for("competency_detail",competency_id=competency_id))
    snap=competency_operational_snapshot(competency_id)
    if snap["counts"]["blocking"] or snap["counts"]["ready"]:
        flash("Ainda existem pendências ou cobranças prontas que não foram enviadas. Resolva-as antes de fechar.","error"); return redirect(url_for("competency_detail",competency_id=competency_id))
    create_db_backup(f"antes_fechar_competencia_{competency_id}")
    conn=db(); conn.execute("UPDATE competencies SET status='FECHADA',closed_at=?,updated_at=? WHERE id=?",(now_iso(),now_iso(),competency_id)); conn.commit(); conn.close()
    audit_event("FECHAR_COMPETENCIA","competencia",competency_id,"Competência fechada e protegida contra alterações estruturais.",competency_id=competency_id)
    flash("Competência fechada. Pagamentos e lembretes continuam disponíveis; alterações estruturais exigem reabertura.","success")
    return redirect(url_for("competency_detail",competency_id=competency_id))


@app.route("/competencies/<int:competency_id>/reopen",methods=["POST"])
def competency_reopen(competency_id):
    comp=get_competency(competency_id)
    if not comp: abort(404)
    create_db_backup(f"antes_reabrir_competencia_{competency_id}")
    conn=db(); conn.execute("UPDATE competencies SET status='PREPARACAO',closed_at=NULL,reopened_at=?,updated_at=? WHERE id=?",(now_iso(),now_iso(),competency_id)); conn.commit(); conn.close()
    recompute_competency_status(competency_id)
    audit_event("REABRIR_COMPETENCIA","competencia",competency_id,"Competência reaberta para ajustes.",competency_id=competency_id)
    flash("Competência reaberta para ajustes.","success")
    return redirect(url_for("competency_detail",competency_id=competency_id))


@app.route("/competencies/<int:competency_id>/simulation")
def competency_simulation(competency_id):
    snap=competency_operational_snapshot(competency_id)
    if not snap: abort(404)
    conn=db()
    try: combined=combined_review_data(conn,competency_id)
    finally: conn.close()
    return render_template('competency_simulation.html',**snap,**combined)


@app.route("/competencies/<int:competency_id>/sync",methods=["POST"])
def competency_sync(competency_id):
    conn = db()
    try:
        conn.execute("BEGIN IMMEDIATE")
        assert_competency_mutable(conn, competency_id)
        sync_competency_companies(conn, competency_id)
        conn.commit()
    except ValueError as exc:
        conn.rollback()
        flash(str(exc), "error")
        return redirect(url_for("competency_detail", competency_id=competency_id))
    finally:
        conn.close()
    audit_event("SINCRONIZAR_EMPRESAS", "competencia", competency_id, "Empresas ativas sincronizadas.", competency_id=competency_id)
    recompute_competency_status(competency_id)
    flash("Novas empresas incluídas. Envios e pagamentos anteriores foram preservados; o próximo envio considera somente pendentes.", "success")
    return redirect(url_for("competency_detail", competency_id=competency_id))


@app.route("/competencies/<int:competency_id>/control",methods=["POST"])
def competency_control(competency_id):
    if not ensure_competency_editable(competency_id): return redirect(url_for("competency_apuration_page",competency_id=competency_id))
    f=request.files.get("control_file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        flash("Envie uma planilha .xlsx de controle.","error"); return redirect(url_for("competency_apuration_page",competency_id=competency_id))
    tmp=CONTROL_DIR/f"upload_{uuid.uuid4().hex}.xlsx"; f.save(tmp)
    try:
        stats,unpriced=process_control_file(competency_id,tmp,f.filename,auto_register_missing=False)
        msg=(f"Controle processado: {stats['a_prazo']} linha(s) A PRAZO; "
             f"{stats['candidate']} complementar(es) candidato(s); {stats['charged']} incluído(s); "
             f"{stats['unpriced']} sem preço; {stats['historical_duplicates']} já cobrado(s) anteriormente; "
             f"{stats['duplicates']} duplicado(s) na própria planilha; {stats['paid_at_exam']} sem cobrança mensal conforme cadastro; "
             f"{stats['other_unit']} atendimento(s) sem empresa identificada; {stats['preserved_companies']} cobrança(s) enviada(s)/paga(s) preservada(s). "
             f"Valor apurado da planilha: {money(stats['charged_total'])}. Valores calculados pela coluna VALOR da planilha.")
        if stats.get("unregistered_exam_counts"):
            exams_txt = ", ".join(f"{name} ({count})" for name, count in sorted(stats["unregistered_exam_counts"].items()))
            msg += f" Exames encontrados sem cadastro ativo: {exams_txt}. Eles foram identificados, mas não cobrados."
        if stats.get("companies_without_email"):
            msg += f" {stats['companies_without_email']} empresa(s) precisam de e-mail de cobrança ou de uma responsável em grupo."
        if stats.get("unmatched_documents"):
            docs_txt = ", ".join(format_document(x) for x in stats["unmatched_documents"][:5])
            msg += f" Não cadastrados encontrados no SETOR: {docs_txt}."
        billable_candidates = max(0, stats["candidate"] - stats["paid_at_exam"])
        category = "warning" if stats["unpriced"] or stats["other_unit"] or stats.get("companies_without_email") or stats.get("unregistered_exam_counts") or (stats["charged"] == 0 and billable_candidates > 0) else "success"
        flash(msg, category)
    except Exception as exc:
        flash(f"Não foi possível processar a planilha: {exc}","error")
    finally:
        try: tmp.unlink()
        except Exception: pass
    return redirect(url_for("competency_apuration_page",competency_id=competency_id))


@app.route("/competencies/<int:competency_id>/reprocess",methods=["POST"])
def competency_reprocess(competency_id):
    comp=ensure_competency_editable(competency_id)
    if not comp: return redirect(url_for("competency_apuration_page",competency_id=competency_id))
    if not comp or not comp["control_stored_name"]:
        flash("Nenhuma planilha de controle armazenada para reprocesar.","error"); return redirect(url_for("competency_apuration_page",competency_id=competency_id))
    p=stored_file_path(comp["control_stored_name"], control=True)
    if not p.exists():
        flash("Arquivo de controle não encontrado no servidor.","error")
    else:
        try:
            stats,_=process_control_file(competency_id,p,comp["control_filename"] or p.name,auto_register_missing=False); flash(f"Recalculado: {stats['a_prazo']} A PRAZO, {stats['preserved_companies']} cobrança(s) preservada(s), {stats['charged']} incluído(s), {stats['unpriced']} sem preço, {stats['historical_duplicates']} já cobrado(s) e {stats['paid_at_exam']} sem cobrança mensal conforme cadastro.","success" if not stats["unpriced"] and not stats["historical_duplicates"] else "warning")
        except Exception as exc: flash(str(exc),"error")
    return redirect(url_for("competency_apuration_page",competency_id=competency_id))


@app.route("/competencies/<int:competency_id>/apuration-audit.xlsx")
def competency_apuration_audit_export(competency_id):
    comp=get_competency(competency_id)
    if not comp: abort(404)
    conn=db()
    run=conn.execute("SELECT * FROM apuration_runs WHERE competency_id=? ORDER BY id DESC LIMIT 1",(competency_id,)).fetchone()
    if not run:
        conn.close(); flash("Ainda não existe processamento para gerar o relatório.","warning"); return redirect(url_for("competency_apuration_page",competency_id=competency_id))
    rows=conn.execute("SELECT a.*,e.source_value FROM apuration_audit a LEFT JOIN exam_items e ON e.competency_id=a.competency_id AND e.source_row=a.source_row AND e.fingerprint=a.fingerprint WHERE a.run_id=? ORDER BY a.source_row,a.id",(run["id"],)).fetchall(); conn.close()
    out=io.BytesIO(); x=Workbook(); ws=x.active; ws.title="RESUMO"
    stats=json.loads(run["stats_json"] or "{}")
    ws.append(["INDICADOR","RESULTADO"]); labels=[("Linhas lidas","rows"),("Última linha real do arquivo","effective_last_row"),("A PRAZO","a_prazo"),("Complementares candidatos","candidate"),("Incluídos com preço","charged"),("Sem preço","unpriced"),("Já cobrados anteriormente","historical_duplicates"),("Duplicados na planilha","duplicates"),("Sem cobrança mensal conforme cadastro","paid_at_exam"),("Cobranças preservadas","preserved_companies"),("Valor preservado","preserved_total"),("Linhas de empresas com cobrança preservada","preserved_rows"),("Empresa não cadastrada","other_unit"),("Tipos de exame sem cadastro ativo","ignored_exam_type")]
    for label,key in labels: ws.append([label,stats.get(key,0) or 0])
    ws.append(["Aba processada", stats.get("sheet_name", "")])
    ws.append(["Fonte dos preços", stats.get("price_source_label", "Coluna VALOR da planilha")])
    ws.append(["Empresas cadastradas neste processamento", stats.get("registered_companies", 0)])
    ws.append(["Empresas com VALOR ausente no controle", stats.get("companies_without_price", 0)])
    ws.append(["Empresas sem e-mail próprio", stats.get("companies_without_email", 0)])
    ws.append(["Valor apurado da planilha", stats.get("charged_total", 0)])
    ws.append(["Soma de VALOR no controle", stats.get("source_total_reference", 0)])
    we = x.create_sheet("EXAMES_DETECTADOS")
    we.append(["TIPO ENCONTRADO NO CONTROLE", "QUANTIDADE", "CADASTRO ATIVO"])
    unknown = stats.get("unregistered_exam_counts", {}) or {}
    for exam_name, count in sorted((stats.get("detected_exam_counts", {}) or {}).items()):
        we.append([exam_name, count, "NÃO" if exam_name in unknown else "SIM"])
    wc = x.create_sheet("EMPRESAS_E_PENDENCIAS")
    wc.append(["EMPRESA", "CNPJ/CPF", "CADASTRADA AGORA", "ATENDIMENTOS", "INCLUIDOS", "SEM PRECO", "E-MAIL PROPRIO", "VALOR APURADO", "PENDENCIA"])
    for item in stats.get("company_summary", []):
        pending = []
        if item.get("unpriced"): pending.append("Preencher VALOR na planilha: " + ", ".join(item.get("unpriced_exams", {})))
        if item.get("missing_email"): pending.append("Cadastrar e-mail ou responsável de grupo")
        wc.append([item["name"], format_document(item["document"]), "SIM" if item.get("registered_now") else "NÃO", item["rows"], item["charged"], item["unpriced"], "AUSENTE" if item.get("missing_email") else "CADASTRADO", item["total"], "; ".join(pending)])
    for item in stats.get("unrecognized_companies", []):
        wc.append([item["name"], format_document(item["document"]), "NÃO", item["rows"], 0, "", "", "", item["reason"]])
    wd=x.create_sheet("LINHAS_ANALISADAS"); hd=["LINHA","DECISÃO","MOTIVO","CNPJ/CPF","EMPRESA","FUNCIONÁRIO","TIPO ORIGINAL","TIPO RECONHECIDO","DATA","VALOR UNITÁRIO DA PLANILHA","VALOR APURADO","VALOR DO CONTROLE","FINGERPRINT"]; wd.append(hd)
    for r in rows: wd.append([r["source_row"],r["decision"],r["reason"],format_document(r["document"]),r["company_name"],r["employee"],r["source_exam"],r["canonical_exam"],r["exam_date"],r["unit_price"] if r["unit_price"] is not None else "",r["amount"] if r["amount"] is not None else "",r["source_value"] if r["source_value"] is not None else "",r["fingerprint"]])
    wi=x.create_sheet("INCLUÍDOS"); wi.append(hd)
    wg=x.create_sheet("IGNORADOS"); wg.append(hd)
    for r in rows:
        target=wi if r["decision"] in {"INCLUÍDO","PENDENTE"} else wg
        target.append([r["source_row"],r["decision"],r["reason"],format_document(r["document"]),r["company_name"],r["employee"],r["source_exam"],r["canonical_exam"],r["exam_date"],r["unit_price"] if r["unit_price"] is not None else "",r["amount"] if r["amount"] is not None else "",r["source_value"] if r["source_value"] is not None else "",r["fingerprint"]])
    for sh in x.worksheets:
        for cell in sh[1]: cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="15314B"); cell.alignment=Alignment(horizontal="center")
        sh.freeze_panes="A2"; sh.auto_filter.ref=sh.dimensions
        for col in range(1,sh.max_column+1): sh.column_dimensions[get_column_letter(col)].width=min(42,max(12,max(len(str(sh.cell(rr,col).value or "")) for rr in range(1,min(sh.max_row,300)+1))+2))
    for sh in [wd,wi,wg]:
        for c in sh["J"][1:]+sh["K"][1:]+sh["L"][1:]: c.number_format='R$ #,##0.00'
    for cell in wc["H"][1:]: cell.number_format='R$ #,##0.00'
    x.save(out); out.seek(0)
    return send_file(out,as_attachment=True,download_name=f"AUDITORIA_APURACAO_{MONTHS[comp['month']]}_{comp['year']}.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/exam-items/<int:item_id>/override-duplicate",methods=["POST"])
def exam_item_override_duplicate(item_id):
    conn = db()
    item = conn.execute("SELECT * FROM exam_items WHERE id=?", (item_id,)).fetchone()
    if not item:
        conn.close(); abort(404)
    try:
        conn.execute("BEGIN IMMEDIATE")
        assert_competency_mutable(conn, item["competency_id"], complementary=True)
        if item["status"] != "JA_COBRADO":
            raise ValueError("Este item não está marcado como já cobrado.")
        value = parse_money_strict(item["source_value"])
        if value is None:
            raise ValueError("Preencha o VALOR na planilha e reprocesse antes de liberar este atendimento.")
        conn.execute("UPDATE exam_items SET status='OK',unit_price=?,total=?,duplicate_override=1,note=? WHERE id=?", (value, value, "Cobrança duplicada liberada manualmente pelo usuário.", item_id))
        totals = conn.execute("SELECT total FROM exam_items WHERE competency_id=? AND company_id=? AND status='OK'", (item["competency_id"], item["company_id"])).fetchall()
        total = sum(int(round(float(r["total"] or 0) * 100)) for r in totals) / 100
        conn.execute("UPDATE competency_companies SET complementary_amount=?,complementary_amount_manual=0,complementary_value_pending=COALESCE((SELECT complementary_value_required FROM companies WHERE id=company_id),0),updated_at=? WHERE competency_id=? AND company_id=?", (total, now_iso(), item["competency_id"], item["company_id"]))
        conn.commit()
    except ValueError as exc:
        conn.rollback()
        flash(str(exc), "error")
        return redirect(url_for("competency_apuration_page", competency_id=item["competency_id"]))
    finally:
        conn.close()
    recompute_competency_status(item["competency_id"])
    audit_event("LIBERAR_DUPLICIDADE", "exam_item", item_id, "Item marcado como já cobrado foi liberado manualmente.", company_id=item["company_id"], competency_id=item["competency_id"])
    flash("Exame liberado para cobrança nesta competência.", "success")
    return redirect(url_for("competency_apuration_page", competency_id=item["competency_id"]))


@app.route("/competencies/<int:competency_id>/apuration.xlsx")
def competency_apuration(competency_id):
    comp=get_competency(competency_id)
    if not comp: abort(404)
    conn=db()
    detail=conn.execute("""SELECT c.name company,c.cnpj,e.employee,e.exam_name,e.exam_date,e.job_title,e.receipt,e.unit_price,e.total,e.status,e.source_document,e.source_company FROM exam_items e JOIN companies c ON c.id=e.company_id WHERE e.competency_id=? ORDER BY c.name,e.employee,e.exam_name""",(competency_id,)).fetchall()
    summary=conn.execute("""SELECT c.name company,c.cnpj,cc.complementary_amount final_total,cc.complementary_amount_manual manual,cc.complementary_value_pending pending,COUNT(CASE WHEN e.status='OK' THEN 1 END) exam_count,SUM(CASE WHEN e.status='OK' THEN COALESCE(e.total,0) ELSE 0 END) total,SUM(CASE WHEN e.status='SEM_PRECO' THEN 1 ELSE 0 END) unpriced,SUM(CASE WHEN e.status='JA_COBRADO' THEN 1 ELSE 0 END) already_billed FROM competency_companies cc JOIN companies c ON c.id=cc.company_id LEFT JOIN exam_items e ON e.competency_id=cc.competency_id AND e.company_id=cc.company_id WHERE cc.competency_id=? GROUP BY c.id ORDER BY c.name""",(competency_id,)).fetchall(); conn.close()
    out=io.BytesIO(); x=Workbook(); ws=x.active; ws.title="RESUMO"; headers=["EMPRESA","CNPJ/CPF","QTD. EXAMES","TOTAL FINAL","SEM VALOR NA PLANILHA","JÁ COBRADO","TOTAL APURADO","VALOR CONFIRMADO MANUALMENTE","CONFERÊNCIA PENDENTE"]; ws.append(headers)
    for r in summary: ws.append([r["company"],format_cnpj(r["cnpj"]),r["exam_count"] or 0,float(r["final_total"] or 0),r["unpriced"] or 0,r["already_billed"] or 0,float(r["total"] or 0),"SIM" if r["manual"] else "NÃO","SIM" if r["pending"] else "NÃO"])
    wd=x.create_sheet("DETALHAMENTO"); hd=["EMPRESA","CNPJ/CPF","COLABORADOR","EXAME","DATA","FUNÇÃO","Nº RECIBO","VALOR UNITÁRIO","TOTAL","STATUS","CNPJ/CPF ORIGEM","EMPRESA ORIGEM"]; wd.append(hd)
    for r in detail: wd.append([r["company"],format_cnpj(r["cnpj"]),r["employee"],r["exam_name"],r["exam_date"] or "",r["job_title"] or "",r["receipt"] or "",r["unit_price"] if r["unit_price"] is not None else "",r["total"] if r["total"] is not None else "",r["status"],format_document(r["source_document"] or r["cnpj"]),r["source_company"] or r["company"]])
    for sh in (ws,wd):
        for cell in sh[1]: cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="15314B"); cell.alignment=Alignment(horizontal="center")
        sh.freeze_panes="A2"; sh.auto_filter.ref=sh.dimensions
        for col in range(1,sh.max_column+1): sh.column_dimensions[get_column_letter(col)].width=min(38,max(12,max(len(str(sh.cell(r,col).value or "")) for r in range(1,min(sh.max_row,200)+1))+2))
    for c in ws["D"][1:]: c.number_format='R$ #,##0.00'
    for c in wd["H"][1:]+wd["I"][1:]: c.number_format='R$ #,##0.00'
    x.save(out); out.seek(0)
    filename=f"APURACAO_COMPLEMENTARES_{MONTHS[comp['month']]}_{comp['year']}.xlsx"
    return send_file(out,as_attachment=True,download_name=filename,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/competencies/<int:competency_id>/control/delete", methods=["POST"])
def competency_control_delete(competency_id):
    if not ensure_competency_editable(competency_id):
        return redirect(url_for("competency_apuration_page", competency_id=competency_id))
    if not confirm_delete_request():
        flash("Digite EXCLUIR para apagar a planilha de controle e a apuração.", "error")
        return redirect(url_for("competency_apuration_page", competency_id=competency_id))
    comp = get_competency(competency_id)
    if not comp:
        abort(404)
    create_db_backup(f"antes_apagar_apuracao_{competency_id}")
    conn = db()
    try:
        conn.execute("BEGIN IMMEDIATE")
        assert_competency_mutable(conn, competency_id, complementary=True)
        conn.execute("DELETE FROM exam_items WHERE competency_id=?", (competency_id,))
        conn.execute("UPDATE competency_companies SET complementary_amount=0,updated_at=? WHERE competency_id=?", (now_iso(), competency_id))
        conn.execute("UPDATE competencies SET control_filename=NULL,control_hash=NULL,control_stored_name=NULL,processed_at=NULL,updated_at=? WHERE id=?", (now_iso(), competency_id))
        conn.commit()
    except ValueError as exc:
        conn.rollback()
        flash(str(exc), "error")
        return redirect(url_for("competency_apuration_page", competency_id=competency_id))
    finally:
        conn.close()
    remove_control_file(comp["control_stored_name"])
    recompute_competency_status(competency_id)
    audit_event("APAGAR_APURACAO", "competencia", competency_id, "Planilha de controle e apuração apagadas.", competency_id=competency_id)
    flash("Planilha de controle e apuração de complementares apagadas.", "success")
    return redirect(url_for("competency_apuration_page", competency_id=competency_id))


@app.route("/exam-items/<int:item_id>/delete", methods=["POST"])
def exam_item_delete(item_id):
    conn = db()
    item = conn.execute("SELECT * FROM exam_items WHERE id=?", (item_id,)).fetchone()
    if not item:
        conn.close(); abort(404)
    try:
        create_db_backup(f"antes_excluir_item_{item_id}")
        conn.execute("BEGIN IMMEDIATE")
        assert_competency_mutable(conn, item["competency_id"], complementary=True)
        conn.execute("DELETE FROM exam_items WHERE id=?", (item_id,))
        totals = conn.execute("SELECT total FROM exam_items WHERE competency_id=? AND company_id=? AND status='OK'", (item["competency_id"], item["company_id"])).fetchall()
        total = sum(int(round(float(r["total"] or 0) * 100)) for r in totals) / 100
        conn.execute("UPDATE competency_companies SET complementary_amount=?,complementary_amount_manual=0,complementary_value_pending=COALESCE((SELECT complementary_value_required FROM companies WHERE id=company_id),0),updated_at=? WHERE competency_id=? AND company_id=?", (total, now_iso(), item["competency_id"], item["company_id"]))
        conn.commit()
    except ValueError as exc:
        conn.rollback()
        flash(str(exc), "error")
        return redirect(url_for("competency_company", competency_id=item["competency_id"], company_id=item["company_id"]))
    finally:
        conn.close()
    recompute_competency_status(item["competency_id"])
    audit_event("EXCLUIR_EXAME_APURACAO", "exam_item", item_id, f"{item['exam_name']} de {item['employee']} removido da apuração.", company_id=item["company_id"], competency_id=item["competency_id"])
    flash("Exame removido da apuração.", "success")
    return redirect(url_for("competency_company", competency_id=item["competency_id"], company_id=item["company_id"]))


@app.route("/competencies/<int:competency_id>/documents/<billing_type>",methods=["POST"])
def competency_documents(competency_id,billing_type):
    if not ensure_competency_editable(competency_id): return redirect(url_for("competency_documents_page",competency_id=competency_id))
    billing_type=billing_type.upper()
    f=request.files.get("zip_file")
    if not f or not f.filename.lower().endswith(".zip"):
        flash("Envie um arquivo ZIP.","error"); return redirect(url_for("competency_documents_page",competency_id=competency_id))
    try:
        added,skipped,unmatched,run_id,import_stats=import_documents_zip(competency_id,billing_type,f)
        msg=f"Arquivos no ZIP: {import_stats['total']}. Importados: {added}. Duplicados: {import_stats['duplicates']}. Ignorados: {import_stats['skipped']}."
        if unmatched:
            preview=" | ".join(unmatched[:5])
            extra=f" (+{len(unmatched)-5} outros)" if len(unmatched)>5 else ""
            msg+=f" Sem empresa identificada: {len(unmatched)}. {preview}{extra}"
        flash(msg,"warning" if unmatched else "success")
    except Exception as exc: flash(f"Falha ao importar ZIP: {exc}","error")
    return redirect(url_for("competency_documents_page",competency_id=competency_id))


@app.route("/competencies/<int:competency_id>/company/<int:company_id>")
def competency_company(competency_id,company_id):
    comp=get_competency(competency_id); company=get_company(company_id)
    if not comp or not company: abort(404)
    conn=db(); cc=conn.execute("SELECT * FROM competency_companies WHERE competency_id=? AND company_id=?",(competency_id,company_id)).fetchone(); items=conn.execute("SELECT * FROM exam_items WHERE competency_id=? AND company_id=? ORDER BY employee,exam_name",(competency_id,company_id)).fetchall(); docs=conn.execute("SELECT * FROM documents WHERE competency_id=? AND company_id=? ORDER BY billing_type,original_name",(competency_id,company_id)).fetchall(); logs=conn.execute("SELECT * FROM email_logs WHERE competency_id=? AND company_id=? ORDER BY id DESC LIMIT 20",(competency_id,company_id)).fetchall(); conn.close()
    if not cc: abort(404)
    return render_template("competency_company.html",comp=comp,company=company,cc=cc,items=items,documents=docs,logs=logs)


@app.route("/competencies/<int:competency_id>/company/<int:company_id>/document/<billing_type>",methods=["POST"])
def manual_document(competency_id, company_id, billing_type):
    if not ensure_competency_editable(competency_id):
        return redirect(url_for("competency_company", competency_id=competency_id, company_id=company_id))
    billing_type = billing_type.upper()
    if billing_type not in {"FIXED", "COMPLEMENTARY"}:
        abort(404)
    f = request.files.get("file")
    if not f or Path(f.filename).suffix.lower() not in ALLOWED_DOC_EXT:
        flash("Envie PDF, PNG, JPG, DOC, DOCX, XLS ou XLSX.", "error")
        return redirect(url_for("competency_company", competency_id=competency_id, company_id=company_id))
    conn = db()
    out = None
    try:
        conn.execute("BEGIN IMMEDIATE")
        assert_competency_mutable(conn, competency_id)
        company = conn.execute("SELECT c.* FROM companies c JOIN competency_companies cc ON cc.company_id=c.id WHERE cc.competency_id=? AND c.id=?", (competency_id, company_id)).fetchone()
        if not company:
            raise ValueError("Empresa não encontrada nesta competência.")
        if billing_type == "COMPLEMENTARY" and not accepts_complementary_documents(conn, company):
            raise ValueError("Esta empresa paga complementares no ato e não recebe cobrança mensal de complementares.")
        data = f.read()
        if not data or len(data) > 30 * 1024 * 1024:
            raise ValueError("O documento deve ter conteúdo e no máximo 30 MB.")
        sha = hashlib.sha256(data).hexdigest()
        out_dir = DOCS_DIR / str(competency_id) / billing_type.lower() / str(company_id)
        out_dir.mkdir(parents=True, exist_ok=True)
        out = out_dir / f"{uuid.uuid4().hex}{Path(f.filename).suffix.lower()}"
        conn.execute("INSERT INTO documents(competency_id,company_id,billing_type,original_name,stored_name,sha256,created_at) VALUES(?,?,?,?,?,?,?)", (competency_id, company_id, billing_type, Path(f.filename).name, str(out.relative_to(DATA_DIR)), sha, now_iso()))
        out.write_bytes(data)
        conn.commit()
    except Exception as exc:
        conn.rollback()
        if out is not None:
            out.unlink(missing_ok=True)
        flash("Este mesmo documento já foi adicionado." if isinstance(exc, sqlite3.IntegrityError) else str(exc), "warning" if isinstance(exc, sqlite3.IntegrityError) else "error")
        return redirect(url_for("competency_company", competency_id=competency_id, company_id=company_id))
    finally:
        conn.close()
    audit_event("ADICIONAR_DOCUMENTO", "documento", None, f"Documento {Path(f.filename).name} adicionado manualmente.", company_id=company_id, competency_id=competency_id, metadata={"billing_type": billing_type, "sha256": sha})
    flash("Documento adicionado.", "success")
    recompute_competency_status(competency_id)
    return redirect(url_for("competency_company", competency_id=competency_id, company_id=company_id))


@app.route("/documents/<int:doc_id>/delete",methods=["POST"])
def document_delete(doc_id):
    conn = db()
    document = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
    if not document:
        conn.close(); abort(404)
    try:
        create_db_backup(f"antes_excluir_documento_{doc_id}")
        conn.execute("BEGIN IMMEDIATE")
        assert_competency_mutable(conn, document["competency_id"])
        conn.execute("DELETE FROM documents WHERE id=?", (doc_id,))
        conn.commit()
    except ValueError as exc:
        conn.rollback()
        flash(str(exc), "error")
        return redirect(url_for("competency_company", competency_id=document["competency_id"], company_id=document["company_id"]))
    finally:
        conn.close()
    remove_document_files([document])
    recompute_competency_status(document["competency_id"])
    audit_event("EXCLUIR_DOCUMENTO", "documento", doc_id, f"Documento {document['original_name']} excluído.", company_id=document["company_id"], competency_id=document["competency_id"])
    flash("Documento excluído.", "success")
    return redirect(url_for("competency_company", competency_id=document["competency_id"], company_id=document["company_id"]))


@app.route("/competencies/<int:competency_id>/payments",methods=["POST"])
def payments(competency_id):
    if not get_competency(competency_id):abort(404)
    company_ids=sorted({int(x) for x in request.form.getlist("company_ids") if str(x).isdigit()})
    action=request.form.get("action")
    if not company_ids: flash("Selecione ao menos uma empresa.","warning"); return redirect(url_for("competency_payments_page",competency_id=competency_id))
    actions={"fixed_paid":("fixed",True),"fixed_unpaid":("fixed",False),"comp_paid":("complementary",True),"comp_unpaid":("complementary",False)}
    if action not in actions:abort(400)
    prefix,paid=actions[action]; conn=db(); marks=",".join("?" for _ in company_ids);stamp=now_iso()
    try:
        conn.execute("BEGIN IMMEDIATE")
        if any(payment_is_sending(conn,competency_id,cid) for cid in company_ids):
            raise ValueError("A cobrança de uma empresa selecionada está sendo transmitida. Aguarde o envio concluir antes de alterar o pagamento.")
        if paid:
            changed=conn.execute(f"UPDATE competency_companies SET {prefix}_paid=1,{prefix}_paid_at=?,updated_at=? WHERE competency_id=? AND company_id IN ({marks}) AND {prefix}_paid=0 AND {prefix}_amount>0",[stamp,stamp,competency_id,*company_ids]).rowcount
        else:
            changed=conn.execute(f"UPDATE competency_companies SET {prefix}_paid=0,{prefix}_paid_at=NULL,updated_at=? WHERE competency_id=? AND company_id IN ({marks}) AND {prefix}_paid=1",[stamp,competency_id,*company_ids]).rowcount
        conn.commit();msg=f"{changed} cobrança(s) {'marcada(s) como paga(s)' if paid else 'reaberta(s)'}. Datas de pagamentos já registrados foram preservadas."
        audit_event("ALTERAR_PAGAMENTO","competencia",competency_id,msg,{"action":action,"company_ids":company_ids,"changed":changed},competency_id=competency_id);flash(msg,"success")
    except ValueError as exc:
        conn.rollback();flash(str(exc),"error")
    finally:conn.close()
    return redirect(url_for("competency_payments_page",competency_id=competency_id))


@app.route("/competencies/<int:competency_id>/amounts/<int:company_id>/<billing_type>", methods=["POST"])
@app.route("/competencies/<int:competency_id>/fixed-amounts/<int:company_id>", methods=["POST"])
def update_competency_fixed_amounts(competency_id, company_id, billing_type="FIXED"):
    if billing_type not in {"FIXED", "COMPLEMENTARY"}: abort(404)
    prefix = "fixed" if billing_type == "FIXED" else "complementary"
    """Ajusta a mensalidade somente da competência atual, sem alterar o cadastro-base."""
    conn = db()
    try:
        conn.execute("BEGIN IMMEDIATE")
        assert_competency_mutable(conn, competency_id)
        comp = conn.execute("SELECT * FROM competencies WHERE id=?", (competency_id,)).fetchone()
        if not comp:
            raise ValueError("Competência não encontrada.")
        ctx = delivery_batches.package_context(conn, competency_id, company_id, billing_type, force=True)
        # A prévia pode representar um grupo explícito ou um lote automático de
        # empresas que compartilham o mesmo destinatário. Permita editar exatamente
        # as empresas pertencentes a esse pacote de cobrança.
        allowed = {int(r["company_id"]): r for r in ctx["all_rows"]}
        submitted = {}
        for key, value in request.form.items():
            if not key.startswith(prefix+"_amount_"):
                continue
            try:
                cid = int(key.rsplit("_", 1)[1])
            except ValueError:
                continue
            if cid not in allowed:
                raise ValueError("Uma empresa enviada não pertence a esta cobrança.")
            amount = parse_money_strict(value)
            if amount is None:
                raise ValueError("Informe um valor para cada empresa exibida.")
            submitted[cid] = amount
        if not submitted:
            amount = parse_money_strict(request.form.get(prefix+"_amount"))
            if amount is None:
                raise ValueError("Informe o valor conferido da cobrança.")
            submitted[company_id] = amount
        for cid, amount in submitted.items():
            cc = conn.execute("SELECT * FROM competency_companies WHERE competency_id=? AND company_id=?", (competency_id, cid)).fetchone()
            if not cc:
                raise ValueError("Empresa não encontrada nesta competência.")
            if cc[prefix+"_sent_at"] or cc[prefix+"_paid"]:
                raise ValueError("Não é possível alterar uma cobrança já enviada ou paga. Reabra/corrija o lançamento antes do novo envio.")
            conn.execute(f"UPDATE competency_companies SET {prefix}_amount=?,{prefix}_amount_manual=1,{prefix}_value_pending=0,updated_at=? WHERE competency_id=? AND company_id=?",
                         (amount, now_iso(), competency_id, cid))
        conn.commit()
        audit_event("CONFERIR_VALOR_COMPETENCIA", "competencia", competency_id,
                    "Valor conferido manualmente antes do envio.",
                    {"company_id": company_id, "values": submitted, "billing_type": billing_type}, competency_id=competency_id)
        flash("Valor(es) confirmado(s). Esta modalidade foi liberada para revisão dos documentos e envio nesta competência.", "success")
    except ValueError as exc:
        conn.rollback()
        flash(str(exc), "error")
    finally:
        conn.close()
    return redirect(url_for("email_preview", competency_id=competency_id, company_id=company_id,
        email_type='COMBINED' if request.form.get('return_combined')=='1' else billing_type))


@app.route("/competencies/<int:competency_id>/preview/<int:company_id>/<email_type>")
def email_preview(competency_id,company_id,email_type):
    email_type=email_type.upper()
    if email_type not in {"FIXED","COMPLEMENTARY","COMBINED","FIXED_REMINDER","COMPLEMENTARY_REMINDER"}: abort(404)
    comp=get_competency(competency_id); company=get_company(company_id)
    if not comp or not company or comp['unit_id'] != company['unit_id']: abort(404)
    conn=db()
    try:
        payload=build_email(conn,competency_id,company_id,email_type)
        sections=[]
        envelope=combined_billing.context(conn,competency_id,company_id)
        if email_type=='COMBINED':
            if not envelope:
                flash('Não há novas cobranças pendentes para esta empresa. Consulte os envios no histórico.','info')
                return redirect(url_for('competency_review',competency_id=competency_id))
            packets=envelope['packets']
        else:
            packets=[delivery_batches.package_context(conn,competency_id,company_id,email_type,force=True)]
        for packet in packets:
            prefix='fixed' if packet['billing']=='FIXED' else 'complementary'
            rows=[dict(r) for r in packet['all_rows'] if r['active'] and not r[prefix+'_sent_at'] and not r[prefix+'_paid'] and (prefix=='fixed' or r['bill_complementaries'])]
            if rows: sections.append(dict(prefix=prefix, rows=rows, owner=packet['representative_id'], billing=packet['billing']))
        problem=billing_send_reason(conn,competency_id,email_type,company_id)
    finally: conn.close()
    if not payload: abort(404)
    cfg=smtp_config()
    return render_template('email_preview.html',comp=comp,company=company,payload=payload,email_type=email_type,
        test_mode=bool(cfg.get('test_mode')),test_email=cfg.get('test_email'),amount_sections=sections,
        combined_available=bool(envelope),send_reason=problem)


@app.route('/competencies/<int:competency_id>/preview/<int:company_id>/<email_type>/complementares.xlsx')
def email_complementary_excel(competency_id,company_id,email_type):
    if email_type not in {'FIXED','COMPLEMENTARY','COMBINED','FIXED_REMINDER','COMPLEMENTARY_REMINDER'}: abort(404)
    comp=get_competency(competency_id); company=get_company(company_id)
    if not comp or not company or comp['unit_id']!=company['unit_id']: abort(404)
    conn=db()
    try: payload=build_email(conn,competency_id,company_id,email_type)
    finally: conn.close()
    if not payload or not payload.get('complementary_report'): abort(404)
    report=payload['complementary_report']
    return send_file(io.BytesIO(complementary_report.workbook_bytes(report)),as_attachment=True,
        download_name=f"COMPLEMENTARES_{report['year']}_{report['month']:02d}.xlsx",mimetype=complementary_report.MIME)


@app.route("/competencies/<int:competency_id>/resend-batch/<billing_type>", methods=["POST"])
def resend_competency_batch(competency_id, billing_type):
    billing_type = billing_type.upper()
    if billing_type not in {"FIXED", "COMPLEMENTARY"}:
        return jsonify({"ok": False, "error": "Tipo de cobrança inválido."}), 400
    comp = get_competency(competency_id)
    if not comp:
        return jsonify({"ok": False, "error": "Competência não encontrada."}), 404
    if competency_is_closed(comp):
        return jsonify({"ok": False, "error": "A competência está fechada. Reabra antes de reenviar cobranças."}), 400
    conn = db()
    try:
        data = resend_review_data(conn, competency_id)[billing_type]
        ids = [item["company_id"] for item in data["ready"]]
        blocked = len(data["blocked"])
    finally:
        conn.close()
    if not ids:
        return jsonify({"ok": False, "error": "Não há cobranças já enviadas e não pagas prontas para reenvio nesta modalidade."}), 400
    try:
        jid, created = create_send_job_for_ids(competency_id, billing_type, ids, force=True)
        audit_event(
            "REENVIAR_COMPETENCIA_LOTE", "competencia", competency_id,
            f"Reenvio em lote {billing_type}: {len(ids)} pacote(s) pronto(s); {blocked} bloqueado(s).",
            {"billing_type": billing_type, "packages": len(ids), "blocked": blocked}, competency_id=competency_id,
        )
        return jsonify({"ok": True, "job_id": jid, "created": created, "blocked": blocked})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/competencies/<int:competency_id>/resend/<int:company_id>/<email_type>",methods=["POST"])
def resend_company(competency_id,company_id,email_type):
    email_type=email_type.upper()
    if email_type not in {"FIXED","COMPLEMENTARY"}: return jsonify({"ok":False,"error":"Tipo inválido."}),400
    conn=db(); cc=conn.execute("SELECT * FROM competency_companies WHERE competency_id=? AND company_id=?",(competency_id,company_id)).fetchone(); conn.close()
    if not cc: return jsonify({"ok":False,"error":"Empresa não encontrada na competência."}),404
    conn=db()
    try:
        group_cfg, owner, participants, _ = grouped_billing.context(conn, competency_id, company_id, email_type)
        delivery_ids=billing_delivery_ids(conn,competency_id,email_type,company_id,force=True)
        sent_field = "fixed_sent_at" if email_type == "FIXED" else "complementary_sent_at"
        previously_sent=bool(delivery_ids and conn.execute(
            f"SELECT 1 FROM competency_companies WHERE competency_id=? AND company_id IN ({','.join('?' for _ in delivery_ids)}) AND {sent_field} IS NOT NULL LIMIT 1",
            (competency_id,*delivery_ids)).fetchone())
        representative=billing_delivery_representative(conn,competency_id,email_type,company_id,force=True)
    finally:
        conn.close()
    if group_cfg and owner != company_id: return jsonify({"ok":False,"error":"Use o reenvio pela empresa responsável do agrupamento."}),400
    if not previously_sent: return jsonify({"ok":False,"error":"A cobrança ainda não foi enviada. Use o envio normal."}),400
    try:
        jid,created=create_send_job_for_ids(competency_id,email_type,[representative],force=True)
        audit_event("REENVIAR_COBRANCA","empresa",company_id,f"Reenvio {email_type} solicitado.",company_id=company_id,competency_id=competency_id)
        return jsonify({"ok":True,"job_id":jid,"created":created})
    except Exception as exc:
        return jsonify({"ok":False,"error":str(exc)}),400


@app.route("/jobs/<job_id>/retry-errors",methods=["POST"])
def retry_send_errors(job_id):
    try:
        jid,created=retry_failed_send_job(job_id)
        return jsonify({"ok":True,"job_id":jid,"created":created})
    except Exception as exc:
        return jsonify({"ok":False,"error":str(exc)}),400


@app.route("/competencies/<int:competency_id>/send/<email_type>",methods=["POST"])
def start_send(competency_id,email_type):
    email_type=email_type.upper()
    if email_type not in {"FIXED","COMPLEMENTARY","COMBINED","FIXED_REMINDER","COMPLEMENTARY_REMINDER"}: abort(404)
    if email_type in {'FIXED','COMPLEMENTARY'}: email_type='COMBINED'
    try:
        jid,created=create_send_job(competency_id,email_type)
        return jsonify({"ok":True,"job_id":jid,"created":created})
    except Exception as exc:
        return jsonify({"ok":False,"error":str(exc)}),400


@app.route("/jobs/<job_id>/resolve/<int:group_id>", methods=["POST"])
def resolve_send_result(job_id, group_id):
    try:
        _send_queue.resolve(job_id, group_id, request.form.get("action", ""), request.form.get("original_test_mode"))
        return jsonify({"ok": True})
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/jobs/<job_id>")
def job_status(job_id):
    conn = db()
    try:
        job = conn.execute("SELECT * FROM send_jobs WHERE id=?", (job_id,)).fetchone()
        if not job:
            return jsonify({"ok": False, "error": "Fila não encontrada."}), 404
        groups = conn.execute(
            """SELECT g.id,g.company_id,g.member_ids_json,g.status,g.error,g.delivery_result_json,c.name FROM send_job_groups g JOIN companies c ON c.id=g.company_id
               WHERE g.job_id=? ORDER BY CASE WHEN g.status IN ('UNCERTAIN','PARTIAL','TEST_PARTIAL') THEN 0 ELSE 1 END,g.id DESC""", (job_id,)
        ).fetchall()
        counts = _send_queue.counts(conn, job_id)
    finally:
        conn.close()
    total = max(1, int(job["total_groups"] or 0))
    percent = max(0, min(100, int(round(int(job["processed_groups"] or 0) / total * 100))))
    done = job["status"] in TERMINAL_JOB_STATES
    retryable = done and counts["error"] > 0
    events = []
    for group in groups:
        if len(events) >= 8 and group["status"] not in {"UNCERTAIN", "PARTIAL", "TEST_PARTIAL"}:
            continue
        status = group["status"]
        result = json.loads(group["delivery_result_json"] or "{}")
        events.append({
            "group_id": group["id"], "status": status, "created_at": "",
            "level": "error" if status in {"ERROR", "UNCERTAIN", "PARTIAL", "TEST_PARTIAL"} else "success" if status in {"SENT", "TEST_SENT"} else "info",
            "message": f"{_send_queue.package_label(group, group['name'])}: {GROUP_LABELS.get(status, status)}" + (f" — {group['error']}" if group["error"] else ""),
            "company_ids": _send_queue.member_ids(group),
            "primary_accepted": result.get("primary_accepted"), "accepted_recipients": result.get("accepted", []),
            "mode_unknown": job["test_mode"] is None,
            "reconciliation_url": url_for("resolve_send_result", job_id=job_id, group_id=group["id"]) if done and status in {"UNCERTAIN", "PARTIAL", "TEST_PARTIAL"} else None,
        })
    return jsonify({
        "ok": True, "percent": percent, "done": done, "job": dict(job), "counts": counts,
        "events": events, "retryable": retryable,
        "retry_url": url_for("retry_send_errors", job_id=job_id) if retryable else None,
    })


@app.route("/competencies/<int:competency_id>/payments-template.xlsx")
def payments_template(competency_id):
    comp=get_competency(competency_id)
    if not comp: abort(404)
    out=io.BytesIO(); x=Workbook(); ws=x.active; ws.title="PAGAMENTOS"
    ws.append(["CNPJ/CPF","TIPO","DATA PAGAMENTO","VALOR"])
    ws.append(["12.345.678/0001-90","MENSALIDADE",date.today().strftime("%d/%m/%Y"),500])
    ws.append(["12.345.678/0001-90","COMPLEMENTARES",date.today().strftime("%d/%m/%Y"),120])
    for cell in ws[1]: cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="15314B")
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for col,width in {"A":22,"B":22,"C":20,"D":16}.items(): ws.column_dimensions[col].width=width
    for c in ws["D"][1:]: c.number_format='R$ #,##0.00'
    x.save(out); out.seek(0)
    return send_file(out,as_attachment=True,download_name=f"MODELO_PAGAMENTOS_{MONTHS[comp['month']]}_{comp['year']}.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/competencies/<int:competency_id>/payments-import",methods=["POST"])
def payments_import(competency_id):
    comp=get_competency(competency_id)
    if not comp: abort(404)
    f=request.files.get("payment_file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        flash("Envie a planilha .xlsx de pagamentos.","error"); return redirect(url_for("competency_payments_page",competency_id=competency_id))
    try:
        result=payment_imports.import_workbook(sys.modules[__name__],competency_id,f.read(),f.filename)
        audit_event("IMPORTAR_PAGAMENTOS","competencia",competency_id,f"{result['updated']} baixa(s) importada(s).",{"run_id":result['run_id'],"rejeitados":result['rejected']},competency_id=competency_id)
        flash(f"Pagamentos importados: {result['updated']}. Linhas rejeitadas: {result['rejected']}. O relatório detalhado está disponível nesta tela.","warning" if result['rejected'] else "success")
    except Exception as exc:
        flash(f"Falha ao importar pagamentos: {exc}","error")
    return redirect(url_for("competency_payments_page",competency_id=competency_id))


@app.route("/payment-imports/<int:run_id>/report.xlsx")
def payment_import_report(run_id):
    conn=db()
    run=conn.execute("SELECT * FROM payment_import_runs WHERE id=?",(run_id,)).fetchone()
    conn.close()
    if not run: abort(404)
    wb=Workbook(); ws=wb.active; ws.title="RELATORIO DE BAIXAS"
    ws.append(["LINHA","CNPJ/CPF","EMPRESA","TIPO","VALOR INFORMADO","DATA PAGAMENTO","RESULTADO","MOTIVO"])
    for item in json.loads(run["report_json"]):
        ws.append([item.get(k,"") for k in ("line","document","company","type","informed_value","paid_date","result","reason")])
    for cell in ws[1]: cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="15314B")
    for letter,width in {"A":10,"B":22,"C":35,"D":22,"E":22,"F":22,"G":22,"H":85}.items(): ws.column_dimensions[letter].width=width
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    out=io.BytesIO(); wb.save(out); wb.close(); out.seek(0)
    return send_file(out,as_attachment=True,download_name=f"RELATORIO_PAGAMENTOS_{run_id}.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/document-imports/<int:run_id>/report.xlsx")
def document_import_report(run_id):
    conn=db(); run=conn.execute("SELECT r.*,cp.month,cp.year FROM document_import_runs r JOIN competencies cp ON cp.id=r.competency_id WHERE r.id=?",(run_id,)).fetchone()
    if not run: conn.close(); abort(404)
    rows=conn.execute("SELECT * FROM document_import_items WHERE run_id=? ORDER BY id",(run_id,)).fetchall(); conn.close()
    out=io.BytesIO(); x=Workbook(); ws=x.active; ws.title="RELATORIO"
    ws.append(["ARQUIVO","EMPRESA","RESULTADO","MOTIVO"])
    for r in rows: ws.append([r["filename"],r["company_name"] or "",r["result"],r["reason"] or ""])
    for cell in ws[1]: cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="15314B")
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for col,width in {"A":52,"B":36,"C":22,"D":48}.items(): ws.column_dimensions[col].width=width
    x.save(out); out.seek(0)
    return send_file(out,as_attachment=True,download_name=f"RELATORIO_DOCUMENTOS_{run['billing_type']}_{MONTHS[run['month']]}_{run['year']}.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/history")
def history():
    conn=db(); rows=conn.execute("""SELECT l.*,c.name company,c.cnpj,cp.month,cp.year,u.name unit_name FROM email_logs l JOIN companies c ON c.id=l.company_id JOIN competencies cp ON cp.id=l.competency_id JOIN units u ON u.id=cp.unit_id ORDER BY l.id DESC LIMIT 500""").fetchall(); conn.close(); return render_template("history.html",logs=rows)


@app.route("/history/<int:log_id>/delete", methods=["POST"])
def history_delete(log_id):
    conn=db(); row=conn.execute("SELECT * FROM email_logs WHERE id=?",(log_id,)).fetchone()
    if not row:
        conn.close(); abort(404)
    create_db_backup(f"antes_excluir_historico_{log_id}")
    conn.execute("DELETE FROM email_logs WHERE id=?",(log_id,)); conn.commit(); conn.close()
    flash("Registro do histórico excluído.","success")
    return redirect(url_for("history"))


@app.route("/history/clear", methods=["POST"])
def history_clear():
    if not confirm_delete_request():
        flash("Digite EXCLUIR para limpar todo o histórico de e-mails.","error"); return redirect(url_for("history"))
    create_db_backup("antes_limpar_historico")
    conn=db(); conn.execute("DELETE FROM email_logs"); conn.commit(); conn.close()
    flash("Histórico de e-mails apagado definitivamente.","success")
    return redirect(url_for("history"))


@app.route("/competencies/<int:competency_id>/delete",methods=["POST"])
def competency_delete(competency_id):
    if request.form.get("confirm_text")!="EXCLUIR":
        flash("Digite EXCLUIR para confirmar.","error"); return redirect(url_for("competency_detail",competency_id=competency_id))
    comp=get_competency(competency_id)
    if not comp: abort(404)
    conn=db(); running=conn.execute("SELECT 1 FROM send_jobs WHERE competency_id=? AND status IN ('QUEUED','RUNNING')",(competency_id,)).fetchone()
    if running: conn.close(); flash("Aguarde o envio em andamento terminar antes de excluir.","error"); return redirect(url_for("competency_detail",competency_id=competency_id))
    stored=conn.execute("SELECT stored_name FROM documents WHERE competency_id=?",(competency_id,)).fetchall(); control=conn.execute("SELECT control_stored_name FROM competencies WHERE id=?",(competency_id,)).fetchone(); unit_id=comp["unit_id"]
    create_db_backup(f"antes_excluir_competencia_{competency_id}")
    conn.execute("DELETE FROM competencies WHERE id=?",(competency_id,)); conn.commit(); conn.close()
    for r in stored:
        try: stored_file_path(r["stored_name"]).unlink(missing_ok=True)
        except Exception: pass
    if control and control["control_stored_name"]:
        try: stored_file_path(control["control_stored_name"], control=True).unlink(missing_ok=True)
        except Exception: pass
    flash("Competência excluída definitivamente.","success"); return redirect(url_for("unit_dashboard",unit_id=unit_id))


@app.route("/search")
def global_search():
    q=str(request.args.get("q") or "").strip()
    companies=[]; competencies=[]
    if q:
        conn=db(); qn=normalize_text(q); qdigits=digits(q)
        rows=conn.execute("""SELECT c.*,u.name unit_name FROM companies c JOIN units u ON u.id=c.unit_id ORDER BY c.name""").fetchall()
        for r in rows:
            hay=normalize_text(f"{r['name']} {r['email'] or ''} {r['email_cc'] or ''} {format_document(r['cnpj'])} {r['cnpj']}")
            if qn in hay or (qdigits and qdigits in digits(r['cnpj'])):
                companies.append(r)
                if len(companies)>=50: break
        cps=conn.execute("SELECT cp.*,u.name unit_name FROM competencies cp JOIN units u ON u.id=cp.unit_id ORDER BY cp.year DESC,cp.month DESC").fetchall()
        for cp in cps:
            label=normalize_text(f"{MONTHS.get(cp['month'],'')} {cp['month']} {cp['year']} {cp['unit_name']} {cp['status'] or ''}")
            if qn in label:
                competencies.append(cp)
                if len(competencies)>=30: break
        conn.close()
    return render_template("search.html",q=q,companies=companies,competencies=competencies)


@app.route("/companies/<int:company_id>/timeline")
def company_timeline(company_id):
    company=get_company(company_id)
    if not company: abort(404)
    conn=db()
    charges=conn.execute("""SELECT cc.*,cp.month,cp.year,cp.status,u.name unit_name FROM competency_companies cc
                            JOIN competencies cp ON cp.id=cc.competency_id JOIN units u ON u.id=cp.unit_id
                            WHERE cc.company_id=? ORDER BY cp.year DESC,cp.month DESC""",(company_id,)).fetchall()
    emails=conn.execute("SELECT * FROM email_logs WHERE company_id=? ORDER BY id DESC LIMIT 100",(company_id,)).fetchall()
    audits=conn.execute("SELECT * FROM audit_logs WHERE company_id=? ORDER BY id DESC LIMIT 100",(company_id,)).fetchall()
    docs=conn.execute("""SELECT d.*,cp.month,cp.year FROM documents d JOIN competencies cp ON cp.id=d.competency_id
                         WHERE d.company_id=? ORDER BY d.id DESC LIMIT 100""",(company_id,)).fetchall()
    conn.close()
    return render_template("company_timeline.html",company=company,charges=charges,emails=emails,audits=audits,documents=docs)


@app.route("/registrations/check")
def registrations_check():
    unit_id=request.args.get("unit_id",type=int)
    conn=db(); units_rows=conn.execute("SELECT * FROM units WHERE active=1 ORDER BY name").fetchall()
    sql="""SELECT c.*,u.name unit_name,(SELECT COUNT(*) FROM complementary_prices p WHERE p.company_id=c.id) price_count,
           EXISTS(SELECT 1 FROM billing_group_members m JOIN billing_groups g ON g.id=m.group_id
                  WHERE m.company_id=c.id AND g.consolidate_fixed=1) fixed_group_member
           FROM companies c JOIN units u ON u.id=c.unit_id WHERE c.active=1"""
    params=[]
    if unit_id: sql+=" AND c.unit_id=?"; params.append(unit_id)
    sql+=" ORDER BY u.name,c.name"
    rows=conn.execute(sql,params).fetchall(); conn.close()
    items=[]; counts={"total":0,"issues":0,"no_email":0,"no_fixed":0,"invalid_doc":0,"no_prices":0}
    for r in rows:
        issues=[]; counts["total"]+=1
        if not valid_email(r["email"]): issues.append("E-mail principal ausente/inválido"); counts["no_email"]+=1
        if float(r["fixed_value"] or 0)<=0 and not r["fixed_group_member"] and not r["fixed_value_required"]:
            issues.append("Mensalidade fixa sem valor"); counts["no_fixed"]+=1
        if not valid_company_document(r["cnpj"]): issues.append("CNPJ/CPF inválido"); counts["invalid_doc"]+=1
        if r["fixed_value_required"] or r["complementary_value_required"]: counts["no_prices"]+=1
        if issues: counts["issues"]+=1
        items.append({"row":r,"issues":issues})
    return render_template("registrations_check.html",items=items,counts=counts,units=units_rows,selected_unit=unit_id)


@app.route("/audit")
def audit_history():
    conn=db(); rows=conn.execute("""SELECT a.*,c.name company_name,cp.month,cp.year,u.name unit_name
                                    FROM audit_logs a LEFT JOIN companies c ON c.id=a.company_id
                                    LEFT JOIN competencies cp ON cp.id=a.competency_id
                                    LEFT JOIN units u ON u.id=cp.unit_id ORDER BY a.id DESC LIMIT 1000""").fetchall(); conn.close()
    return render_template("audit.html",rows=rows)


@app.route("/settings/backup",methods=["POST"])
def settings_backup_create():
    path=create_db_backup("manual")
    audit_event("BACKUP_MANUAL","sistema",description=f"Backup manual criado: {path.name}")
    flash(f"Backup criado: {path.name}","success")
    return redirect(url_for("settings"))


@app.route("/settings/backup/<path:name>/download")
def settings_backup_download(name):
    safe=Path(name).name; path=BACKUP_DIR/safe
    if not path.exists() or path.suffix.lower() not in {".db",".zip"}: abort(404)
    return send_file(path,as_attachment=True,download_name=safe,mimetype="application/octet-stream")


@app.route("/settings/backup/<path:name>/restore",methods=["POST"])
def settings_backup_restore(name):
    if normalize_text(request.form.get("confirm_text"))!="RESTAURAR":
        flash("Digite RESTAURAR para confirmar a restauração do banco.","error"); return redirect(url_for("settings"))
    safe=Path(name).name; source=BACKUP_DIR/safe
    if not source.exists() or source.suffix.lower()!=".db": abort(404)
    try:
        current,has_files=data_backups.restore(sys.modules[__name__],source)
        audit_event("RESTAURAR_BACKUP","sistema",description=f"Banco restaurado de {safe}. Backup anterior: {current.name}")
        flash(f"Backup {safe} restaurado. "+("Documentos e controles também restaurados." if has_files else "Este backup antigo contém somente o banco; os arquivos precisam existir na pasta data."),"success" if has_files else "warning")
    except Exception as exc:
        flash(f"Não foi possível restaurar: {exc}","error")
    return redirect(url_for("settings"))


@app.route("/settings", methods=["GET", "POST"])
def settings():
    if request.method == "POST":
        setting_set("smtp_host", "smtp.gmail.com")
        security = request.form.get("smtp_security", "starttls")
        if security not in {"starttls", "ssl"}:
            security = "starttls"
        setting_set("smtp_security", security)
        setting_set("smtp_port", "465" if security == "ssl" else "587")
        setting_set("smtp_username", request.form.get("smtp_username", "").strip())
        setting_set("sender_name", request.form.get("sender_name", "EDGE Saúde Ocupacional").strip())
        setting_set("sender_email", request.form.get("sender_email", "").strip())
        setting_set("test_email", request.form.get("test_email", "").strip())
        setting_set("test_mode", "1" if request.form.get("test_mode") else "0")
        setting_set("email_signature", request.form.get("email_signature", "EDGE Saúde Ocupacional").strip())
        password = request.form.get("smtp_password", "")
        if password:
            setting_set("smtp_password", encrypt_secret(password.replace(" ", "")))
        flash("Configurações do Gmail salvas.", "success")
        return redirect(url_for("settings"))
    cfg = smtp_config()
    backups=[]
    for bp in sorted(BACKUP_DIR.glob("*.db"),key=lambda x:x.stat().st_mtime,reverse=True)[:30]:
        backups.append({"name":bp.name,"bundle_name":bp.with_suffix('.zip').name if bp.with_suffix('.zip').is_file() else None,"size":bp.stat().st_size,"modified":datetime.fromtimestamp(bp.stat().st_mtime).strftime("%d/%m/%Y %H:%M")})
    return render_template("settings.html", cfg=cfg, has_password=bool(setting_get("smtp_password")), signature=cfg["email_signature"],backups=backups)


@app.route("/settings/test-email", methods=["POST"])
def settings_test_email():
    cfg = smtp_config()
    dest = cfg["test_email"] or cfg["sender_email"]
    try:
        smtp_send(dest, "", "TESTE DE ENVIO - EDGE COBRANÇAS", "<p>Configuração do Gmail funcionando.</p><p><strong>EDGE Saúde Ocupacional</strong></p>", "Configuração do Gmail funcionando.\nEDGE Saúde Ocupacional")
        flash("E-mail de teste enviado com sucesso.", "success")
    except Exception as exc:
        flash(f"Falha no teste: {exc}", "error")
    return redirect(url_for("settings"))


@app.route("/change-password", methods=["POST"])
def change_password():
    current = request.form.get("current_password", "")
    new = request.form.get("new_password", "")
    h = setting_get("admin_password_hash")
    if not h or not check_password_hash(h, current):
        flash("Senha atual incorreta.", "error")
    elif len(new) < 8:
        flash("A nova senha deve ter pelo menos 8 caracteres.", "error")
    else:
        setting_set("admin_password_hash", generate_password_hash(new))
        flash("Senha alterada.", "success")
    return redirect(url_for("settings"))


grouped_billing.register_routes(sys.modules[__name__])


@app.route("/healthz")
def healthz():
    return {"ok": True, "module": "envio-cobrancas", "version": APP_VERSION}



recover_send_jobs()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=os.environ.get("FLASK_DEBUG") == "1")
