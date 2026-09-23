from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify, session, abort
import os, re, tempfile, unicodedata, zipfile, io, sqlite3, shutil, subprocess, logging, secrets, hmac, json, hashlib
from io import BytesIO
from pathlib import Path
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from copy import deepcopy

import pandas as pd
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from docxtpl import DocxTemplate, RichText
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from lxml import etree

try:
    import psycopg2
    import psycopg2.extras
except Exception:
    psycopg2 = None

app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_CONTENT_LENGTH", 250 * 1024 * 1024))
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("SESSION_COOKIE_SECURE", "1") not in {"0", "false", "False"}
app.permanent_session_lifetime = timedelta(hours=int(os.environ.get("SESSION_HOURS", "8")))

logging.basicConfig(level=os.environ.get("LOG_LEVEL", "INFO"), format="%(asctime)s %(levelname)s %(name)s: %(message)s")
logger = logging.getLogger("site_interno_edge")

APP_TITLE = "Sistema Interno"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "ENCAMINHAMENTOS PERIODICO MCP.docx")
ALLOWED_EXTENSIONS = {".xls", ".xlsx", ".html", ".htm"}
EXCEL_EXTENSIONS = {".xls", ".xlsx"}
RELATORIOS_EXTENSIONS = {".xls", ".xlsx", ".zip"}
RENUM_ALLOWED_EXTENSIONS = {".docx", ".zip"}
ESOCIAL_ALLOWED_EXTENSIONS = {".xls", ".xlsx"}
ESOCIAL_MONTHS = {
    "JANEIRO": "JANEIRO", "FEVEREIRO": "FEVEREIRO", "MARCO": "MARÇO",
    "ABRIL": "ABRIL", "MAIO": "MAIO", "JUNHO": "JUNHO",
    "JULHO": "JULHO", "AGOSTO": "AGOSTO", "SETEMBRO": "SETEMBRO",
    "OUTUBRO": "OUTUBRO", "NOVEMBRO": "NOVEMBRO", "DEZEMBRO": "DEZEMBRO",
}
FISICO_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "ATESTADO_FISICO_MENTAL_TEMPLATE.docx")
ATESTADO_MEDICO_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "ATESTADO_MEDICO_TEMPLATE.docx")
PCD_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "MODELO LAUDO PCD.docx")
ENCAMINHAMENTO_PREENCHIMENTO_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "ENCAMINHAMENTO_PREENCHIMENTO_TEMPLATE.docx")
ENCAMINHAMENTO_COMPLEMENTARES_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "ENCAMINHAMENTO_COMPLEMENTARES_TEMPLATE.docx")
ANAMNESE_OCUPACIONAL_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "ANAMNESE_OCUPACIONAL_TEMPLATE.docx")
ASO_MANUAL_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "ASO_MANUAL_TEMPLATE.docx")

# Dados oficiais das unidades usados em todos os modelos da Clínica.
# Manter centralizado evita divergência de endereço/telefone entre documentos.
CLINIC_LOCATIONS = {
    'belem': {
        'label': 'BELÉM, PA',
        'cidade_uf': 'BELÉM, PA',
        'cidade_data': 'BELÉM, PA',
        'cidade_hifen': 'BELÉM-PA',
        'endereco': 'TRAVESSA DO CHACO, Nº2546, ENTRE ALMIRANTE BARROSO E JOÃO PAULO – BELÉM – PA',
        'telefone': '91– 3349-6948',
        'medico_fisico_mental': 'CRM Nº 4480 – RQE Nº6041 PA',
        'cnpj_edge_aso': '28.589.436/0001-87',
        'info_doutor_aso': 'Nº 4480 – RQE Nº 6041 PA',
        'uf': 'PA',
    },
    'macapa': {
        'label': 'MACAPÁ, AP',
        'cidade_uf': 'MACAPÁ, AP',
        'cidade_data': 'MACAPÁ, AP',
        'cidade_hifen': 'MACAPÁ-AP',
        'endereco': 'RUA ELIÉZER LEVY, Nº 2583, TREM, MACAPÁ-AP',
        'telefone': '91– 98356-8044',
        'medico_fisico_mental': 'CRM Nº 002800 – RQE Nº959 AP',
        'cnpj_edge_aso': '33.789.248/0001-32',
        'info_doutor_aso': 'Nº 0002800 – RQE Nº 959 AP',
        'uf': 'AP',
    },
}

EXAMES_A_PRAZO_SOLO_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "exames_a_prazo_templates", "MODELO_PRA_EMPRESA_SOLO.xlsx")
EXAMES_A_PRAZO_GROUP_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "exames_a_prazo_templates", "MODELO_PARA_MULTIPLAS_EMPRESAS.xlsx")
EXAMES_A_PRAZO_SOURCE_EXTENSIONS = {".xlsx"}
EXAMES_A_PRAZO_REQUEST_EXTENSIONS = {".xlsx", ".zip"}
EXAMES_A_PRAZO_MAX_REQUEST_FILES = 200
EXAMES_A_PRAZO_MAX_SINGLE_XLSX_BYTES = int(os.environ.get("EXAMES_A_PRAZO_MAX_SINGLE_XLSX_MB", "150")) * 1024 * 1024
EXAMES_A_PRAZO_MAX_ZIP_UNCOMPRESSED_BYTES = int(os.environ.get("EXAMES_A_PRAZO_MAX_ZIP_MB", "600")) * 1024 * 1024
BASE_DIR = os.path.dirname(__file__)
DATA_DIR = os.environ.get("RENDER_DISK_PATH") or os.environ.get("DATA_DIR") or BASE_DIR
os.makedirs(DATA_DIR, exist_ok=True)
ESOCIAL_BASE_CACHE_DIR = os.path.join(DATA_DIR, "esocial_base_sessions")
os.makedirs(ESOCIAL_BASE_CACHE_DIR, exist_ok=True)
RELATORIOS_EMPRESAS_CNPJ_PATH = os.path.join(DATA_DIR, "relatorios_empresas_cnpj.json")
RELATORIOS_EMPRESAS_CNPJ_EXTENSIONS = {".xls", ".xlsx"}
FISICO_DB_PATH = os.path.join(DATA_DIR, "fisico_mental.db")
AUTH_DB_PATH = os.path.join(DATA_DIR, "usuarios.db")
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
USE_POSTGRES = bool(DATABASE_URL)
USER_ROLES = {"admin": "Administrador", "supervisor": "Supervisor", "operador": "Operador", "visualizador": "Visualizador"}
ROLE_DESCRIPTIONS = {
    "admin": "Acesso total, usuários, auditoria e configurações.",
    "supervisor": "Usa as funcionalidades e consulta auditoria, sem gerenciar usuários.",
    "operador": "Usa as funcionalidades operacionais do sistema.",
    "visualizador": "Acesso de leitura às telas, sem processar, gerar ou alterar dados."
}

# =========================
# UTILIDADES GERAIS
# =========================
def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text)

def sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]+', "_", str(name))
    name = re.sub(r"\s+", " ", name).strip()
    return name or "arquivo"



def get_max_upload_mb() -> int:
    return max(1, int(app.config.get("MAX_CONTENT_LENGTH", 250 * 1024 * 1024)) // (1024 * 1024))

def validate_uploaded_file(file_storage, allowed_extensions: set[str], label: str = "o arquivo") -> tuple[bool, str]:
    """Valida upload antes de processar, com mensagens seguras para o usuário.

    A validação por extensão não substitui antivírus/sandbox, mas reduz risco operacional
    e evita que fluxos internos tentem abrir formatos inesperados.
    """
    if not file_storage or not getattr(file_storage, "filename", ""):
        return False, f"Selecione {label}."
    original_name = Path(file_storage.filename).name
    suffix = Path(original_name).suffix.lower()
    if suffix not in allowed_extensions:
        permitidos = ", ".join(sorted(allowed_extensions))
        return False, f"Formato inválido em {original_name}. Envie apenas: {permitidos}."

    # Quando o navegador informa tamanho, bloqueia arquivos vazios e reforça limite configurado.
    content_length = getattr(file_storage, "content_length", None) or 0
    max_bytes = int(app.config.get("MAX_CONTENT_LENGTH", 250 * 1024 * 1024))
    if content_length and content_length > max_bytes:
        return False, f"{original_name} excede o limite de {get_max_upload_mb()} MB."

    # Proteção contra nomes maliciosos/path traversal.
    safe_name = secure_filename(original_name)
    if not safe_name or safe_name in {".", ".."}:
        return False, f"Nome de arquivo inválido em {original_name}."
    return True, ""

def unique_path(path: str) -> str:
    base, ext = os.path.splitext(path)
    if not os.path.exists(path):
        return path
    i = 2
    while True:
        new = f"{base} ({i}){ext}"
        if not os.path.exists(new):
            return new
        i += 1

def safe_extract_zip(zip_path: Path, destination: Path) -> None:
    """Extrai ZIP sem permitir caminhos maliciosos fora da pasta destino."""
    destination = destination.resolve()
    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        for member in zip_ref.infolist():
            member_name = member.filename.replace("\\", "/")
            if not member_name or member_name.startswith("/"):
                raise ValueError("ZIP contém caminho absoluto inválido.")
            target = (destination / member_name).resolve()
            if destination not in target.parents and target != destination:
                raise ValueError("ZIP contém caminho inseguro e foi bloqueado.")
            if member.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with zip_ref.open(member) as source, open(target, "wb") as dest:
                shutil.copyfileobj(source, dest)

def auth_get_conn():
    if USE_POSTGRES:
        if psycopg2 is None:
            raise RuntimeError("psycopg2-binary não está instalado. Adicione ao requirements.txt.")
        return psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    conn = sqlite3.connect(AUTH_DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def db_param() -> str:
    return "%s" if USE_POSTGRES else "?"

def db_id_type() -> str:
    return "SERIAL PRIMARY KEY" if USE_POSTGRES else "INTEGER PRIMARY KEY AUTOINCREMENT"

def row_get(row, key, default=None):
    if row is None:
        return default
    try:
        return row[key]
    except Exception:
        return default

def db_fetchone(conn, sql, params=()):
    cur = conn.cursor()
    cur.execute(sql, params)
    return cur.fetchone()

def db_fetchall(conn, sql, params=()):
    cur = conn.cursor()
    cur.execute(sql, params)
    return cur.fetchall()

def db_execute(conn, sql, params=()):
    cur = conn.cursor()
    cur.execute(sql, params)
    return cur

def validate_password_strength(password: str) -> tuple[bool, str]:
    if not password or len(password) < 8:
        return False, 'A senha deve ter pelo menos 8 caracteres.'
    if not re.search(r'[A-Za-z]', password) or not re.search(r'\d', password):
        return False, 'A senha deve conter letras e números.'
    return True, ''

def table_columns(conn, table_name: str) -> set[str]:
    try:
        if USE_POSTGRES:
            rows = db_fetchall(conn, """
                SELECT column_name FROM information_schema.columns
                WHERE table_schema = 'public' AND table_name = %s
            """, (table_name,))
        else:
            rows = db_fetchall(conn, f"PRAGMA table_info({table_name})")
        cols = set()
        for row in rows:
            if USE_POSTGRES:
                cols.add(str(row_get(row, 'column_name')))
            else:
                cols.add(str(row_get(row, 'name')))
        return cols
    except Exception:
        return set()

def ensure_column(conn, table_name: str, column_name: str, definition: str) -> None:
    if column_name in table_columns(conn, table_name):
        return
    db_execute(conn, f"ALTER TABLE {table_name} ADD COLUMN {column_name} {definition}")

def init_auth_db():
    with auth_get_conn() as conn:
        db_execute(conn, f"""CREATE TABLE IF NOT EXISTS usuarios (
            id {db_id_type()},
            nome TEXT NOT NULL,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            cargo TEXT NOT NULL DEFAULT 'operador',
            ativo INTEGER NOT NULL DEFAULT 1,
            criado_em TEXT NOT NULL,
            atualizado_em TEXT NOT NULL,
            ultimo_login TEXT,
            precisa_trocar_senha INTEGER NOT NULL DEFAULT 0
        )""")
        db_execute(conn, f"""CREATE TABLE IF NOT EXISTS audit_logs (
            id {db_id_type()},
            user_id INTEGER,
            username TEXT,
            cargo TEXT,
            acao TEXT NOT NULL,
            detalhe TEXT,
            ip TEXT,
            user_agent TEXT,
            criado_em TEXT NOT NULL
        )""")
        conn.commit()

    with auth_get_conn() as conn:
        ensure_column(conn, 'usuarios', 'ultimo_login', 'TEXT')
        ensure_column(conn, 'usuarios', 'precisa_trocar_senha', 'INTEGER NOT NULL DEFAULT 0')
        conn.commit()

    bootstrap_admin_from_env()

def normalize_role(cargo: str) -> str:
    cargo = (cargo or 'operador').strip().lower()
    return cargo if cargo in USER_ROLES else 'operador'

def auth_user_count() -> int:
    with auth_get_conn() as conn:
        return int(row_get(db_fetchone(conn, 'SELECT COUNT(*) AS total FROM usuarios'), 'total', 0))

def auth_get_user_by_username(username: str):
    with auth_get_conn() as conn:
        return db_fetchone(conn, f'SELECT * FROM usuarios WHERE lower(username) = lower({db_param()})', (username.strip(),))

def auth_get_user_by_id(user_id):
    if not user_id:
        return None
    with auth_get_conn() as conn:
        return db_fetchone(conn, f'SELECT * FROM usuarios WHERE id = {db_param()}', (int(user_id),))

def auth_current_user():
    return auth_get_user_by_id(session.get('user_id'))

def auth_is_logged_in() -> bool:
    user = auth_current_user()
    return bool(user and int(row_get(user, 'ativo', 0)) == 1)

def auth_is_admin() -> bool:
    user = auth_current_user()
    return bool(user and int(row_get(user, 'ativo', 0)) == 1 and row_get(user, 'cargo') == 'admin')

def current_user_role() -> str:
    user = auth_current_user()
    return row_get(user, 'cargo', 'visualizador') if user else 'anonimo'

def has_role(*roles: str) -> bool:
    return current_user_role() in set(roles)

def is_readonly_user() -> bool:
    return current_user_role() == 'visualizador'

def audit_log(acao: str, detalhe: str = '') -> None:
    try:
        user = auth_current_user()
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with auth_get_conn() as conn:
            db_execute(conn, f'''INSERT INTO audit_logs (user_id, username, cargo, acao, detalhe, ip, user_agent, criado_em)
                VALUES ({db_param()}, {db_param()}, {db_param()}, {db_param()}, {db_param()}, {db_param()}, {db_param()}, {db_param()})''', (
                row_get(user, 'id') if user else None,
                row_get(user, 'username') if user else 'sistema',
                row_get(user, 'cargo') if user else 'sistema',
                acao[:80],
                (detalhe or '')[:1000],
                (request.headers.get('X-Forwarded-For') or request.remote_addr or '')[:120] if request else '',
                (request.headers.get('User-Agent') or '')[:250] if request else '',
                now,
            ))
            conn.commit()
    except Exception as exc:
        logger.warning("Falha ao gravar auditoria: %s", exc)

def auth_create_user(nome: str, username: str, password: str, cargo: str, precisa_trocar_senha: int = 1) -> tuple[bool, str]:
    nome = (nome or '').strip()
    username = (username or '').strip()
    cargo = normalize_role(cargo)
    if not nome or len(nome) < 2:
        return False, 'Informe o nome completo do usuário.'
    if not re.fullmatch(r'[A-Za-z0-9._-]{3,40}', username or ''):
        return False, 'O usuário deve ter de 3 a 40 caracteres e usar apenas letras, números, ponto, traço ou underline.'
    ok_pwd, msg_pwd = validate_password_strength(password)
    if not ok_pwd:
        return False, msg_pwd
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        with auth_get_conn() as conn:
            db_execute(conn, f'''INSERT INTO usuarios (nome, username, password_hash, cargo, ativo, criado_em, atualizado_em, precisa_trocar_senha)
                VALUES ({db_param()}, {db_param()}, {db_param()}, {db_param()}, 1, {db_param()}, {db_param()}, {db_param()})''', (nome, username, generate_password_hash(password), cargo, now, now, int(precisa_trocar_senha)))
            conn.commit()
        audit_log('usuario_criado', f'Usuário {username} criado com cargo {cargo}.')
        return True, 'Usuário criado com sucesso.'
    except Exception as exc:
        if 'unique' in str(exc).lower() or 'duplicate' in str(exc).lower():
            return False, 'Já existe um usuário com esse login.'
        logger.exception("Erro ao criar usuário")
        return False, 'Não foi possível criar o usuário.'

def auth_list_users():
    with auth_get_conn() as conn:
        return db_fetchall(conn, 'SELECT id, nome, username, cargo, ativo, criado_em, atualizado_em, ultimo_login, precisa_trocar_senha FROM usuarios ORDER BY nome')

def auth_update_user(user_id: int, nome: str, username: str, cargo: str, ativo: int, password: str = '', precisa_trocar_senha: int = 0) -> tuple[bool, str]:
    user = auth_get_user_by_id(user_id)
    if not user:
        return False, 'Usuário não encontrado.'
    nome = (nome or '').strip()
    username = (username or '').strip()
    cargo = normalize_role(cargo)
    ativo = 1 if str(ativo) == '1' else 0
    if not nome or len(nome) < 2:
        return False, 'Informe o nome do usuário.'
    if not re.fullmatch(r'[A-Za-z0-9._-]{3,40}', username or ''):
        return False, 'Login inválido.'
    if password:
        ok_pwd, msg_pwd = validate_password_strength(password)
        if not ok_pwd:
            return False, msg_pwd
    if row_get(user, 'cargo') == 'admin' and (ativo == 0 or cargo != 'admin'):
        with auth_get_conn() as conn:
            admins_ativos = row_get(db_fetchone(conn, f"SELECT COUNT(*) AS total FROM usuarios WHERE cargo = 'admin' AND ativo = 1 AND id <> {db_param()}", (user_id,)), 'total', 0)
        if int(admins_ativos) == 0:
            return False, 'Não é permitido remover ou desativar o único administrador ativo.'
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        with auth_get_conn() as conn:
            if password:
                db_execute(conn, f'''UPDATE usuarios SET nome = {db_param()}, username = {db_param()}, cargo = {db_param()}, ativo = {db_param()}, password_hash = {db_param()}, precisa_trocar_senha = {db_param()}, atualizado_em = {db_param()} WHERE id = {db_param()}''', (nome, username, cargo, ativo, generate_password_hash(password), int(precisa_trocar_senha), now, user_id))
            else:
                db_execute(conn, f'''UPDATE usuarios SET nome = {db_param()}, username = {db_param()}, cargo = {db_param()}, ativo = {db_param()}, precisa_trocar_senha = {db_param()}, atualizado_em = {db_param()} WHERE id = {db_param()}''', (nome, username, cargo, ativo, int(precisa_trocar_senha), now, user_id))
            conn.commit()
        audit_log('usuario_atualizado', f'Usuário {username} atualizado. Cargo={cargo}, ativo={ativo}.')
        return True, 'Usuário atualizado com sucesso.'
    except Exception as exc:
        if 'unique' in str(exc).lower() or 'duplicate' in str(exc).lower():
            return False, 'Já existe outro usuário com esse login.'
        logger.exception("Erro ao atualizar usuário")
        return False, 'Não foi possível atualizar o usuário.'

def auth_delete_user(user_id: int) -> tuple[bool, str]:
    user = auth_get_user_by_id(user_id)
    if not user:
        return False, 'Usuário não encontrado.'
    if int(row_get(user, 'id')) == int(session.get('user_id') or 0):
        return False, 'Você não pode excluir o próprio usuário logado.'
    if row_get(user, 'cargo') == 'admin':
        with auth_get_conn() as conn:
            admins_ativos = row_get(db_fetchone(conn, f"SELECT COUNT(*) AS total FROM usuarios WHERE cargo = 'admin' AND ativo = 1 AND id <> {db_param()}", (user_id,)), 'total', 0)
        if int(admins_ativos) == 0:
            return False, 'Não é permitido excluir o único administrador ativo.'
    username = row_get(user, 'username')
    with auth_get_conn() as conn:
        db_execute(conn, f'DELETE FROM usuarios WHERE id = {db_param()}', (user_id,))
        conn.commit()
    audit_log('usuario_excluido', f'Usuário {username} excluído.')
    return True, 'Usuário excluído com sucesso.'

def bootstrap_admin_from_env():
    username = os.environ.get('ADMIN_USERNAME', '').strip()
    password = os.environ.get('ADMIN_PASSWORD', '')
    nome = os.environ.get('ADMIN_NAME', 'Administrador').strip() or 'Administrador'
    if not username or not password or auth_get_user_by_username(username):
        return
    auth_create_user(nome, username, password, 'admin', precisa_trocar_senha=0)

def auth_mark_login(user_id: int):
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with auth_get_conn() as conn:
        db_execute(conn, f'UPDATE usuarios SET ultimo_login = {db_param()} WHERE id = {db_param()}', (now, user_id))
        conn.commit()

def auth_change_own_password(user_id: int, old_password: str, new_password: str, confirm: str) -> tuple[bool, str]:
    user = auth_get_user_by_id(user_id)
    if not user or not check_password_hash(row_get(user, 'password_hash'), old_password):
        return False, 'Senha atual inválida.'
    if new_password != confirm:
        return False, 'As novas senhas não conferem.'
    ok_pwd, msg_pwd = validate_password_strength(new_password)
    if not ok_pwd:
        return False, msg_pwd
    with auth_get_conn() as conn:
        db_execute(conn, f'UPDATE usuarios SET password_hash = {db_param()}, precisa_trocar_senha = 0, atualizado_em = {db_param()} WHERE id = {db_param()}', (generate_password_hash(new_password), datetime.now().strftime('%Y-%m-%d %H:%M:%S'), user_id))
        conn.commit()
    audit_log('senha_alterada', 'Usuário alterou a própria senha.')
    return True, 'Senha alterada com sucesso.'

def get_recent_audit_logs(limit: int = 150):
    with auth_get_conn() as conn:
        return db_fetchall(conn, f'SELECT username, cargo, acao, detalhe, ip, criado_em FROM audit_logs ORDER BY id DESC LIMIT {int(limit)}')

def generate_csrf_token():
    token = session.get('_csrf_token')
    if not token:
        token = secrets.token_urlsafe(32)
        session['_csrf_token'] = token
    return token

def validate_csrf():
    if request.method != 'POST':
        return None
    if request.endpoint in {'login', 'setup_admin'}:
        return None
    session_token = session.get('_csrf_token')
    received = request.form.get('_csrf_token') or request.headers.get('X-CSRF-Token')
    if not session_token or not received or not hmac.compare_digest(str(session_token), str(received)):
        audit_log('csrf_bloqueado', f'Endpoint bloqueado: {request.endpoint}')
        abort(403)
    return None

def admin_required():
    if not auth_is_admin():
        flash('Apenas administradores podem acessar esta área.', 'error')
        return redirect(url_for('home'))
    return None

def supervisor_or_admin_required():
    if not has_role('admin', 'supervisor'):
        flash('Apenas administradores ou supervisores podem acessar esta área.', 'error')
        return redirect(url_for('home'))
    return None

@app.context_processor
def inject_auth_context():
    return {'current_user': auth_current_user(), 'is_admin': auth_is_admin(), 'user_roles': USER_ROLES, 'role_descriptions': ROLE_DESCRIPTIONS, 'csrf_token': generate_csrf_token}


@app.before_request
def attach_request_id():
    request.request_id = secrets.token_hex(8)
    return None

@app.after_request
def apply_security_headers(response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    if request.is_secure or os.environ.get("RENDER"):
        response.headers.setdefault("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
    return response

@app.before_request
def require_login():
    endpoint = request.endpoint or ''
    public_endpoints = {'login', 'setup_admin', 'healthz'}
    if endpoint.startswith('static'):
        return None
    if auth_user_count() == 0:
        if endpoint != 'setup_admin':
            return redirect(url_for('setup_admin'))
        return None
    if endpoint in public_endpoints:
        return None
    if not auth_is_logged_in():
        return redirect(url_for('login', next=request.path))
    validate_csrf()
    session.permanent = True
    user = auth_current_user()
    if row_get(user, 'precisa_trocar_senha', 0) and endpoint not in {'minha_conta', 'logout'}:
        flash('Por segurança, altere sua senha antes de continuar.', 'warning')
        return redirect(url_for('minha_conta'))
    if request.method == 'POST' and is_readonly_user():
        audit_log('post_bloqueado_visualizador', f'Endpoint: {endpoint}')
        flash('Seu cargo é Visualizador. Você pode consultar telas, mas não pode gerar, cadastrar, editar ou excluir dados.', 'error')
        return redirect(request.referrer or url_for('home'))
    if request.method == 'POST':
        audit_log('acao_post', f'Endpoint executado: {endpoint}')
    return None

@app.route('/primeiro-acesso', methods=['GET', 'POST'])
def setup_admin():
    if auth_user_count() > 0:
        return redirect(url_for('login'))
    if request.method == 'POST':
        nome = request.form.get('nome', '')
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        confirm = request.form.get('confirm_password', '')
        if password != confirm:
            flash('As senhas não conferem.', 'error')
            return render_template('setup_admin.html', title='Criar administrador')
        ok, msg = auth_create_user(nome, username, password, 'admin', precisa_trocar_senha=0)
        flash(msg, 'success' if ok else 'error')
        if ok:
            user = auth_get_user_by_username(username)
            session.clear()
            session.permanent = True
            session['user_id'] = int(row_get(user, 'id'))
            session['username'] = row_get(user, 'username')
            session['cargo'] = row_get(user, 'cargo')
            auth_mark_login(int(row_get(user, 'id')))
            audit_log('primeiro_admin_criado', f'Administrador inicial {username} criado.')
            return redirect(url_for('home'))
    return render_template('setup_admin.html', title='Criar administrador')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if auth_user_count() == 0:
        return redirect(url_for('setup_admin'))
    if auth_is_logged_in():
        return redirect(url_for('home'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        user = auth_get_user_by_username(username)
        if user and int(row_get(user, 'ativo', 0)) == 1 and check_password_hash(row_get(user, 'password_hash'), password):
            session.clear()
            session.permanent = True
            session['user_id'] = int(row_get(user, 'id'))
            session['username'] = row_get(user, 'username')
            session['cargo'] = row_get(user, 'cargo')
            auth_mark_login(int(row_get(user, 'id')))
            audit_log('login_sucesso', 'Login realizado.')
            return redirect(request.args.get('next') or url_for('home'))
        audit_log('login_falhou', f'Tentativa para usuário: {username}')
        flash('Usuário ou senha inválidos.', 'error')
    return render_template('login.html', title='Login')

@app.route('/logout')
def logout():
    audit_log('logout', 'Usuário saiu do sistema.')
    session.clear()
    return redirect(url_for('login'))

@app.route('/minha-conta', methods=['GET', 'POST'])
def minha_conta():
    if request.method == 'POST':
        ok, msg = auth_change_own_password(int(session.get('user_id')), request.form.get('old_password', ''), request.form.get('new_password', ''), request.form.get('confirm_password', ''))
        flash(msg, 'success' if ok else 'error')
        if ok:
            return redirect(url_for('home'))
    return render_template('minha_conta.html', title='Minha conta')

@app.route('/auditoria', methods=['GET'])
def auditoria():
    block = supervisor_or_admin_required()
    if block:
        return block
    return render_template('auditoria.html', title='Auditoria', logs=get_recent_audit_logs())

@app.route('/usuarios', methods=['GET'])
def usuarios():
    block = admin_required()
    if block:
        return block
    return render_template('usuarios.html', title='Usuários', usuarios=auth_list_users(), roles=USER_ROLES)

@app.route('/usuarios/criar', methods=['POST'])
def usuarios_criar():
    block = admin_required()
    if block:
        return block
    ok, msg = auth_create_user(request.form.get('nome', ''), request.form.get('username', ''), request.form.get('password', ''), request.form.get('cargo', 'operador'), int(request.form.get('precisa_trocar_senha', '1')))
    flash(msg, 'success' if ok else 'error')
    return redirect(url_for('usuarios'))

@app.route('/usuarios/editar', methods=['POST'])
def usuarios_editar():
    block = admin_required()
    if block:
        return block
    user_id = request.form.get('id', '')
    if not str(user_id).isdigit():
        flash('Usuário inválido.', 'error')
        return redirect(url_for('usuarios'))
    ok, msg = auth_update_user(int(user_id), request.form.get('nome', ''), request.form.get('username', ''), request.form.get('cargo', 'operador'), int(request.form.get('ativo', '0')), request.form.get('password', ''), int(request.form.get('precisa_trocar_senha', '0')))
    flash(msg, 'success' if ok else 'error')
    return redirect(url_for('usuarios'))

@app.route('/usuarios/excluir', methods=['POST'])
def usuarios_excluir():
    block = admin_required()
    if block:
        return block
    user_id = request.form.get('id', '')
    if not str(user_id).isdigit():
        flash('Usuário inválido.', 'error')
        return redirect(url_for('usuarios'))
    ok, msg = auth_delete_user(int(user_id))
    flash(msg, 'success' if ok else 'error')
    return redirect(url_for('usuarios'))

init_auth_db()

def encontrar_coluna(df, candidatos, obrigatoria=False):
    cols = list(df.columns)
    normalizadas = {str(c).strip().lower(): c for c in cols}
    for nome in candidatos:
        key = nome.strip().lower()
        if key in normalizadas:
            return normalizadas[key]
    for c in cols:
        c_norm = str(c).strip().lower()
        for nome in candidatos:
            if nome.strip().lower() in c_norm:
                return c
    if obrigatoria:
        raise ValueError(f"Coluna não encontrada. Esperado um destes nomes: {', '.join(candidatos)}")
    return None

def limpar_nome_arquivo(nome_arquivo):
    nome = os.path.splitext(os.path.basename(str(nome_arquivo)))[0]
    nome = re.sub(r"\(\d+\)$", "", nome).strip()
    nome = re.sub(r"\s{2,}", " ", nome)
    return nome

def limpar_nome_pasta_arquivo(nome):
    nome = str(nome or "")
    for c in r'\/:*?"<>|':
        nome = nome.replace(c, "")
    nome = re.sub(r"\s{2,}", " ", nome).strip()
    return nome

def somente_numeros(texto):
    return re.sub(r"\D", "", str(texto or ""))

def formatar_documento(doc):
    numeros = somente_numeros(doc)
    if len(numeros) == 14:
        return f"{numeros[:2]}.{numeros[2:5]}.{numeros[5:8]}/{numeros[8:12]}-{numeros[12:]}"
    if len(numeros) == 11:
        return f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"
    return ""

def extrair_documento_do_final_do_arquivo(nome_arquivo):
    nome = limpar_nome_arquivo(nome_arquivo)
    partes = [p.strip() for p in nome.split(" - ") if p.strip()]
    candidatos = []
    if partes:
        candidatos.append(partes[-1])
    m = re.search(r'([0-9.\-\/]+)\s*$', nome)
    if m:
        candidatos.append(m.group(1).strip())
    for cand in candidatos:
        numeros = somente_numeros(cand)
        if len(numeros) == 14:
            return formatar_documento(numeros)
        if len(numeros) == 11:
            return formatar_documento(numeros)
        encontrados = re.findall(r'\d+', cand)
        if encontrados:
            juntos = "".join(encontrados)
            if len(juntos) >= 14:
                return formatar_documento(juntos[-14:])
            if len(juntos) >= 11:
                return formatar_documento(juntos[-11:])
    nums = re.findall(r'\d+', nome)
    if nums:
        juntos = "".join(nums)
        if len(juntos) >= 14:
            return formatar_documento(juntos[-14:])
        if len(juntos) >= 11:
            return formatar_documento(juntos[-11:])
    return ""

def nome_mes(m):
    meses = ["", "JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO",
             "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"]
    return meses[m]

class UploadedMemoryFile(BytesIO):
    def __init__(self, filename, data):
        super().__init__(data)
        self.filename = filename

def extrair_planilhas_relatorios_uploads(files):
    """Recebe uploads .xls/.xlsx ou .zip e retorna planilhas em memória.

    A funcionalidade de Relatórios normalmente recebe muitos arquivos de
    Convocação. Para facilitar o uso, o usuário também pode enviar o ZIP do mês
    inteiro; o sistema extrai somente .xls/.xlsx e ignora pastas/outros formatos.
    """
    planilhas = []
    for file in files:
        if not file or not getattr(file, "filename", ""):
            continue
        ok, msg = validate_uploaded_file(file, RELATORIOS_EXTENSIONS, "as planilhas")
        if not ok:
            raise ValueError(msg)

        filename = Path(file.filename).name
        suffix = Path(filename).suffix.lower()
        dados = file.read()

        if suffix == ".zip":
            try:
                with zipfile.ZipFile(BytesIO(dados), "r") as zip_ref:
                    for member in zip_ref.infolist():
                        member_name = member.filename.replace("\\", "/")
                        if member.is_dir() or not member_name:
                            continue
                        member_path = Path(member_name)
                        if member_path.suffix.lower() not in EXCEL_EXTENSIONS:
                            continue
                        if member_path.is_absolute() or ".." in member_path.parts:
                            raise ValueError("O ZIP contém caminho inseguro e foi bloqueado.")
                        with zip_ref.open(member) as source:
                            planilhas.append((member_path.name, source.read()))
            except zipfile.BadZipFile as exc:
                raise ValueError(f"Arquivo ZIP inválido: {filename}.") from exc
        else:
            planilhas.append((filename, dados))
    return planilhas

def salvar_planilhas_relatorios_uploads(files, destination: Path):
    destination.mkdir(parents=True, exist_ok=True)
    planilhas = extrair_planilhas_relatorios_uploads(files)
    saved = []
    for index, (filename, data) in enumerate(planilhas, start=1):
        safe_name = secure_filename(Path(filename).name)
        if not safe_name:
            safe_name = f"planilha_{index}.xlsx"
        path = destination / f"{index:03d}_{safe_name}"
        path.write_bytes(data)
        saved.append(path)
    return saved

# =========================
# RELATÓRIOS DE PERIÓDICOS + BASE DO MÊS + COMPARAÇÃO
# =========================
def _texto_upper(valor) -> str:
    """Texto padronizado em maiúsculo para os relatórios."""
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass
    if isinstance(valor, (datetime, pd.Timestamp)):
        return valor.strftime("%d/%m/%Y")
    texto = str(valor).replace("\xa0", " ").strip()
    texto = re.sub(r"\s+", " ", texto)
    if texto.lower() in {"nan", "nat", "none"}:
        return ""
    return texto.upper()


def _normalizar_chave(valor) -> str:
    """Normaliza nomes/cargos para comparação e remoção de duplicidades."""
    texto = _texto_upper(valor)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"[^A-Z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def _formatar_data_br(valor) -> str:
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, pd.Timestamp):
        if pd.isna(valor):
            return ""
        return valor.strftime("%d/%m/%Y")
    texto = str(valor).strip()
    if not texto:
        return ""
    match = re.search(r"(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})", texto)
    if match:
        dia, mes, ano = match.groups()
        ano = f"20{ano}" if len(ano) == 2 else ano
        return f"{int(dia):02d}/{int(mes):02d}/{ano}"
    parsed = pd.to_datetime(valor, dayfirst=True, errors="coerce")
    if not pd.isna(parsed):
        return parsed.strftime("%d/%m/%Y")
    return _texto_upper(texto)


def _coluna_por_candidatos_df(df, candidatos, obrigatoria=False):
    return encontrar_coluna(df, candidatos, obrigatoria=obrigatoria)


def nome_empresa_da_planilha(row, col_empresa, col_setor, nome_arquivo):
    if col_empresa:
        valor = row[col_empresa]
        if not pd.isna(valor) and str(valor).strip():
            return _texto_upper(valor)
    if col_setor:
        valor = row[col_setor]
        if not pd.isna(valor) and str(valor).strip():
            return _texto_upper(valor)
    return _texto_upper(limpar_nome_arquivo(nome_arquivo))


def extrair_data_admissao(row, col_admissao):
    """Retorna exclusivamente a data da coluna ADMISSAO para a função Relatórios."""
    if not col_admissao or col_admissao not in row.index:
        return pd.NaT
    valor = row.get(col_admissao)
    if pd.isna(valor):
        return pd.NaT
    data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
    if pd.isna(data):
        return pd.NaT
    return data


def filtrar_periodicos_do_mes(df, mes, col_nome, col_admissao):
    """Filtra colaboradores pelo mês da ADMISSAO.

    Esta função foi refeita para usar SEMPRE a coluna ADMISSAO.
    VALIDADE, VENCIMENTO, PERIODICO e ULTIMO EXAME não participam do filtro.
    """
    df = df.copy()
    df["_data_periodico_calculada"] = df.apply(
        lambda row: extrair_data_admissao(row, col_admissao),
        axis=1,
    )
    df["_data_periodico_calculada"] = pd.to_datetime(df["_data_periodico_calculada"], errors="coerce")
    df = df.dropna(subset=["_data_periodico_calculada", col_nome])
    return df[df["_data_periodico_calculada"].dt.month == mes].copy()


def _deduplicar_relatorio_por_nome_cargo(df, col_nome, col_cargo):
    if df.empty:
        return df
    df = df.copy()
    df["_nome_dedup"] = df[col_nome].map(_normalizar_chave)
    if col_cargo:
        df["_cargo_dedup"] = df[col_cargo].map(_normalizar_chave)
    else:
        df["_cargo_dedup"] = ""
    df = df.drop_duplicates(subset=["_nome_dedup", "_cargo_dedup"], keep="first")
    return df.drop(columns=["_nome_dedup", "_cargo_dedup"], errors="ignore")


def _documento_digits_relatorios(valor) -> str:
    """Extrai números de CPF/CNPJ preservando melhor valores vindos do Excel."""
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass

    if isinstance(valor, int):
        texto = str(valor)
    elif isinstance(valor, float):
        texto = str(int(valor)) if valor.is_integer() else f"{valor:.0f}"
    else:
        texto = str(valor).strip()
        if re.fullmatch(r"\d+\.0+", texto):
            texto = texto.split(".", 1)[0]
        elif "e" in texto.lower():
            try:
                texto = str(int(float(texto)))
            except Exception:
                pass

    numeros = somente_numeros(texto)
    if len(numeros) == 15 and numeros.endswith("0") and ".0" in str(valor):
        numeros = numeros[:-1]
    if len(numeros) >= 14:
        return numeros[-14:]
    if len(numeros) in {12, 13}:
        # Quando o Excel salva CNPJ como número, zeros à esquerda podem ser perdidos.
        # Mantemos o valor para a rotina de CNPJ completar com zero à esquerda.
        return numeros
    if len(numeros) == 11:
        return numeros
    return numeros


def _cnpj_digits_relatorios(valor) -> str:
    numeros = _documento_digits_relatorios(valor)
    if len(numeros) == 14:
        return numeros
    if len(numeros) in {12, 13}:
        return numeros.zfill(14)
    return ""


def _formatar_cnpj_relatorios(valor) -> str:
    numeros = _cnpj_digits_relatorios(valor)
    return formatar_documento(numeros) if numeros else ""


def carregar_empresas_cnpj_relatorios():
    """Carrega o cadastro persistente de CNPJ -> nome oficial da empresa."""
    try:
        path = Path(RELATORIOS_EMPRESAS_CNPJ_PATH)
        if not path.exists():
            return {}
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {}
        normalizado = {}
        for chave, item in data.items():
            cnpj_digits = _cnpj_digits_relatorios(chave)
            if not cnpj_digits or not isinstance(item, dict):
                continue
            empresa = _texto_upper(item.get("empresa"))
            if not empresa:
                continue
            normalizado[cnpj_digits] = {
                "empresa": empresa,
                "cnpj": formatar_documento(cnpj_digits),
                "updated_at": item.get("updated_at", ""),
            }
        return normalizado
    except Exception:
        logger.exception("Erro ao carregar cadastro de empresas por CNPJ")
        return {}


def salvar_empresas_cnpj_relatorios(cadastros: dict):
    path = Path(RELATORIOS_EMPRESAS_CNPJ_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(cadastros, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


def listar_empresas_cnpj_relatorios(limit=15):
    cadastros = carregar_empresas_cnpj_relatorios()
    itens = sorted(cadastros.values(), key=lambda item: item.get("empresa", ""))
    return {"total": len(itens), "itens": itens[:limit]}


def importar_empresas_cnpj_relatorios(file_storage):
    ok, msg = validate_uploaded_file(file_storage, RELATORIOS_EMPRESAS_CNPJ_EXTENSIONS, "a planilha de empresas/CNPJ")
    if not ok:
        raise ValueError(msg)

    file_storage.seek(0)
    dados = file_storage.read()
    if not dados:
        raise ValueError("A planilha de empresas/CNPJ está vazia.")

    cadastros = carregar_empresas_cnpj_relatorios()
    total_importados = 0
    total_linhas_validas = 0

    try:
        planilhas = pd.read_excel(BytesIO(dados), sheet_name=None)
    except Exception as exc:
        raise ValueError("Não foi possível ler a planilha. Envie um arquivo .xls ou .xlsx válido.") from exc

    for _guia, df in planilhas.items():
        if df is None or df.empty:
            continue
        col_empresa = encontrar_coluna(df, [
            "empresa", "nome da empresa", "razao social", "razão social",
            "nome oficial", "cliente", "nome"
        ])
        col_cnpj = encontrar_coluna(df, ["cnpj", "cnpj da empresa", "documento", "cpf/cnpj"])
        if not col_empresa or not col_cnpj:
            continue
        for _, row in df.iterrows():
            cnpj_digits = _cnpj_digits_relatorios(row.get(col_cnpj))
            empresa = _texto_upper(row.get(col_empresa))
            if not cnpj_digits or not empresa:
                continue
            total_linhas_validas += 1
            cadastros[cnpj_digits] = {
                "empresa": empresa,
                "cnpj": formatar_documento(cnpj_digits),
                "updated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
            }
            total_importados += 1

    if total_linhas_validas == 0:
        raise ValueError("Nenhuma empresa foi cadastrada. A planilha precisa ter colunas de EMPRESA/NOME DA EMPRESA e CNPJ.")

    salvar_empresas_cnpj_relatorios(cadastros)
    return total_importados


def obter_empresa_cnpj_relatorios(row, col_cnpj, col_empresa, col_setor, nome_arquivo, cadastros=None):
    """Define empresa oficial e CNPJ formatado para Relatório e Base do Mês."""
    cadastros = cadastros if cadastros is not None else carregar_empresas_cnpj_relatorios()
    cnpj_digits = ""
    if col_cnpj and col_cnpj in row.index:
        cnpj_digits = _cnpj_digits_relatorios(row.get(col_cnpj))
    if not cnpj_digits:
        cnpj_digits = _cnpj_digits_relatorios(extrair_documento_do_final_do_arquivo(nome_arquivo))

    if cnpj_digits and cnpj_digits in cadastros:
        item = cadastros[cnpj_digits]
        return item.get("empresa", ""), item.get("cnpj", formatar_documento(cnpj_digits))

    empresa = nome_empresa_da_planilha(row, col_empresa, col_setor, nome_arquivo)
    cnpj_formatado = formatar_documento(cnpj_digits) if cnpj_digits else ""
    return empresa, cnpj_formatado


def titulo_empresa_relatorios(empresa, cnpj):
    empresa = _texto_upper(empresa)
    cnpj = _texto_upper(cnpj)
    return f"{empresa} - {cnpj}" if cnpj else empresa


def _ler_convocacao(file):
    file.seek(0)
    df = pd.read_excel(file)
    col_nome = _coluna_por_candidatos_df(df, ["nome", "funcionario", "funcionário"], obrigatoria=True)
    col_admissao = _coluna_por_candidatos_df(
        df,
        ["admissao", "admissão", "data admissao", "data admissão", "data de admissao", "data de admissão"],
        obrigatoria=True,
    )
    col_cargo = _coluna_por_candidatos_df(df, ["cargo", "função", "funcao"])
    col_empresa = _coluna_por_candidatos_df(df, ["empresa", "razão social", "razao social"])
    col_setor = _coluna_por_candidatos_df(df, ["setor", "ges"])
    col_comp = _coluna_por_candidatos_df(df, ["complementares", "complementar", "exames_obg", "exames obrigatorios", "exames obrigatórios"])
    col_cnpj = _coluna_por_candidatos_df(df, ["cnpj", "cnpj da empresa", "documento da empresa", "cpf/cnpj"])
    return df, col_nome, col_admissao, col_cargo, col_empresa, col_setor, col_comp, col_cnpj


def criar_relatorio(files, mes, cadastros_empresas=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório"

    cor_empresa = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal="center", vertical="center")
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    linha = 1
    cadastros_empresas = cadastros_empresas if cadastros_empresas is not None else carregar_empresas_cnpj_relatorios()

    def escrever_bloco(titulo, registros):
        nonlocal linha
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
        cell = ws.cell(row=linha, column=1, value=titulo)
        cell.font = Font(size=13, bold=True)
        cell.fill = cor_empresa
        cell.alignment = left_wrap
        for col in range(1, 4):
            ws.cell(row=linha, column=col).border = borda
        ws.row_dimensions[linha].height = 35
        linha += 1

        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
        cell = ws.cell(row=linha, column=1, value="NOME DO FUNCIONÁRIO")
        cell.font = Font(bold=True)
        cell.alignment = center
        for col in range(1, 4):
            ws.cell(row=linha, column=col).border = borda
        linha += 1

        if registros is None or registros.empty:
            ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
            cell = ws.cell(row=linha, column=1, value=f"NÃO HÁ COLABORADORES COM ADMISSÃO PARA O MÊS DE {nome_mes(mes)}")
            cell.font = Font(color="FF0000", bold=True)
            cell.alignment = center
            for col in range(1, 4):
                ws.cell(row=linha, column=col).border = borda
            linha += 3
            return

        for _, row_item in registros.iterrows():
            nome = _texto_upper(row_item[col_nome])
            ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
            cell = ws.cell(row=linha, column=1, value=nome)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            for col in range(1, 4):
                ws.cell(row=linha, column=col).border = borda
            linha += 1
        linha += 2

    for file in files:
        try:
            df, col_nome, col_admissao, col_cargo, col_empresa, col_setor, _col_comp, col_cnpj = _ler_convocacao(file)
            filtrado = filtrar_periodicos_do_mes(df, mes, col_nome, col_admissao)
            filtrado = _deduplicar_relatorio_por_nome_cargo(filtrado, col_nome, col_cargo)

            if filtrado.empty:
                if not df.empty:
                    empresa, cnpj = obter_empresa_cnpj_relatorios(df.iloc[0], col_cnpj, col_empresa, col_setor, file.filename, cadastros_empresas)
                    titulo = titulo_empresa_relatorios(empresa, cnpj)
                else:
                    titulo = _texto_upper(limpar_nome_arquivo(file.filename))
                escrever_bloco(titulo, filtrado)
            else:
                filtrado = filtrado.copy()
                titulos = []
                for _, row in filtrado.iterrows():
                    empresa, cnpj = obter_empresa_cnpj_relatorios(row, col_cnpj, col_empresa, col_setor, file.filename, cadastros_empresas)
                    titulos.append(titulo_empresa_relatorios(empresa, cnpj))
                filtrado["_titulo_relatorio_empresa"] = titulos
                for titulo, grupo in filtrado.groupby("_titulo_relatorio_empresa", sort=False):
                    escrever_bloco(titulo, grupo.drop(columns=["_titulo_relatorio_empresa"], errors="ignore"))
        except Exception as e:
            titulo = _texto_upper(limpar_nome_arquivo(file.filename))
            ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
            cell = ws.cell(row=linha, column=1, value=titulo)
            cell.font = Font(size=13, bold=True)
            cell.fill = cor_empresa
            cell.alignment = left_wrap
            for col in range(1, 4):
                ws.cell(row=linha, column=col).border = borda
            ws.row_dimensions[linha].height = 35
            linha += 1
            ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
            cell = ws.cell(row=linha, column=1, value="NOME DO FUNCIONÁRIO")
            cell.font = Font(bold=True)
            cell.alignment = center
            for col in range(1, 4):
                ws.cell(row=linha, column=col).border = borda
            linha += 1
            ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
            cell = ws.cell(row=linha, column=1, value=f"ERRO AO PROCESSAR ARQUIVO: {_texto_upper(e)}")
            cell.font = Font(color="FF0000", bold=True)
            cell.alignment = center
            for col in range(1, 4):
                ws.cell(row=linha, column=col).border = borda
            linha += 3

    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    ws.freeze_panes = "A1"
    return wb


def criar_base(files, mes, cadastros_empresas=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Base do Mês"

    borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    cabecalho_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    cabecalho_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")

    headers = ["EMPRESA", "CNPJ", "NOME", "CARGO", "COMPLEMENTARES"]
    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=idx, value=h)
        cell.font = cabecalho_font
        cell.fill = cabecalho_fill
        cell.alignment = center
        cell.border = borda

    linha = 2
    cadastros_empresas = cadastros_empresas if cadastros_empresas is not None else carregar_empresas_cnpj_relatorios()

    for file in files:
        try:
            df, col_nome, col_admissao, col_cargo, col_empresa, col_setor, col_comp, col_cnpj = _ler_convocacao(file)
            filtrado = filtrar_periodicos_do_mes(df, mes, col_nome, col_admissao)

            for _, row in filtrado.iterrows():
                empresa, documento = obter_empresa_cnpj_relatorios(row, col_cnpj, col_empresa, col_setor, file.filename, cadastros_empresas)
                nome = _texto_upper(row[col_nome])
                cargo = _texto_upper(row[col_cargo]) if col_cargo else ""
                complementares = _texto_upper(row[col_comp]) if col_comp else ""

                valores = [empresa, documento, nome, cargo, complementares]
                for idx, valor in enumerate(valores, start=1):
                    cell = ws.cell(row=linha, column=idx, value=valor if valor is not None else "")
                    cell.border = borda
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    if idx == 2:
                        cell.number_format = "@"
                linha += 1
        except Exception:
            logger.exception("Erro ao processar base do mês: %s", getattr(file, "filename", "arquivo"))
            continue

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 38
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 42
    ws.freeze_panes = "A2"
    return wb


# -------------------------
# COMPARAÇÃO DO RELATÓRIO COM PLANILHA DE CONTROLE DE ASO
# -------------------------
def _salvar_upload_comparacao(file_storage, destination: Path, allowed_extensions: set[str], label: str) -> Path:
    ok, msg = validate_uploaded_file(file_storage, allowed_extensions, label)
    if not ok:
        raise ValueError(msg)
    destination.mkdir(parents=True, exist_ok=True)
    filename = secure_filename(Path(file_storage.filename).name)
    path = destination / filename
    file_storage.save(path)
    return path


def _safe_comparacao_token(token: str) -> str:
    token = str(token or "")
    if not re.fullmatch(r"[A-Za-z0-9_\-]{8,80}", token):
        raise ValueError("Sessão de comparação inválida. Envie os arquivos novamente.")
    return token


def listar_guias_planilha_controle(caminho_controle: Path):
    wb = load_workbook(caminho_controle, read_only=True, data_only=True)
    return list(wb.sheetnames)


def _mapear_header_linha(values):
    mapa = {}
    for idx, value in enumerate(values, start=1):
        chave = _normalizar_chave(value)
        if chave:
            mapa[chave] = idx
    return mapa


def _encontrar_coluna_header(header_map, candidatos):
    candidatos_norm = [_normalizar_chave(c) for c in candidatos]
    for cand in candidatos_norm:
        if cand in header_map:
            return header_map[cand]
    for header, idx in header_map.items():
        for cand in candidatos_norm:
            if cand and cand in header:
                return idx
    return None


def _encontrar_header_controle(ws, max_linhas=25):
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_linhas, values_only=True), start=1):
        header_map = _mapear_header_linha(row)
        col_nome = _encontrar_coluna_header(header_map, ["FUNCIONARIO", "FUNCIONÁRIO", "NOME", "NOME DO FUNCIONARIO", "NOME DO FUNCIONÁRIO"])
        col_data = _encontrar_coluna_header(header_map, ["DATA", "DATA DO EXAME", "DIA"])
        col_tipo = _encontrar_coluna_header(header_map, ["TIPO DE EXAME", "TIPO EXAME", "EXAME", "TIPO"])
        if col_nome and (col_data or col_tipo):
            col_empresa = _encontrar_coluna_header(header_map, ["EMPRESA", "SETOR", "CLIENTE", "RAZAO SOCIAL", "RAZÃO SOCIAL"])
            return row_idx, col_nome, col_data, col_tipo, col_empresa
    return None, None, None, None, None


def criar_indice_controle_asos(caminho_controle: Path, guias_selecionadas):
    """Retorna {NOME_NORMALIZADO: ["DATA - TIPO - EMPRESA", ...]}."""
    indice = {}
    wb = load_workbook(caminho_controle, read_only=True, data_only=True)
    guias_validas = [g for g in guias_selecionadas if g in wb.sheetnames]
    if not guias_validas:
        raise ValueError("Selecione ao menos uma guia válida da planilha de controle.")

    for guia in guias_validas:
        ws = wb[guia]
        header_row, col_nome, col_data, col_tipo, col_empresa = _encontrar_header_controle(ws)
        if not header_row or not col_nome:
            continue
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            nome = row[col_nome - 1] if col_nome <= len(row) else None
            nome_key = _normalizar_chave(nome)
            if not nome_key or nome_key.startswith("EXAMES"):
                continue
            data_txt = _formatar_data_br(row[col_data - 1] if col_data and col_data <= len(row) else "")
            tipo_txt = _texto_upper(row[col_tipo - 1] if col_tipo and col_tipo <= len(row) else "")
            empresa_txt = _texto_upper(row[col_empresa - 1] if col_empresa and col_empresa <= len(row) else "")
            partes = [p for p in [data_txt, tipo_txt, empresa_txt] if p]
            if not partes:
                continue
            info = " - ".join(partes)
            indice.setdefault(nome_key, [])
            if info not in indice[nome_key]:
                indice[nome_key].append(info)
    return indice


def comparar_relatorio_com_controle(caminho_relatorio: Path, caminho_controle: Path, guias_selecionadas, caminho_saida: Path):
    indice = criar_indice_controle_asos(caminho_controle, guias_selecionadas)
    wb = load_workbook(caminho_relatorio)
    ws = wb.active

    borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    header_font = Font(bold=True)
    wrap_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # A coluna D fica ao lado do bloco de nomes do relatório original, que usa A:C mesclado.
    ws.column_dimensions["D"].width = 75
    dentro_lista = False
    linhas_comparadas = 0
    encontrados = 0

    for row_idx in range(1, ws.max_row + 1):
        nome_cell = ws.cell(row=row_idx, column=1)
        valor = _texto_upper(nome_cell.value)
        valor_norm = _normalizar_chave(valor)

        if valor_norm == _normalizar_chave("NOME DO FUNCIONÁRIO"):
            dentro_lista = True
            cell = ws.cell(row=row_idx, column=4, value="DATA / TIPO DE EXAME / EMPRESA")
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = borda
            continue

        if not valor_norm:
            dentro_lista = False
            continue

        if dentro_lista:
            if valor_norm.startswith("NAO HA COLABORADORES") or valor_norm.startswith("ERRO AO PROCESSAR"):
                continue
            linhas_comparadas += 1
            infos = indice.get(valor_norm, [])
            info_cell = ws.cell(row=row_idx, column=4)
            info_cell.value = " | ".join(infos) if infos else ""
            info_cell.alignment = wrap_left
            info_cell.border = borda
            if infos:
                encontrados += 1

    # Pequeno resumo técnico abaixo da última linha, sem alterar os dados do relatório.
    resumo_linha = ws.max_row + 2
    ws.cell(row=resumo_linha, column=1, value="RESUMO DA COMPARAÇÃO")
    ws.cell(row=resumo_linha, column=1).font = Font(bold=True)
    ws.cell(row=resumo_linha + 1, column=1, value=f"NOMES COMPARADOS: {linhas_comparadas}")
    ws.cell(row=resumo_linha + 2, column=1, value=f"NOMES ENCONTRADOS: {encontrados}")
    ws.cell(row=resumo_linha + 3, column=1, value=f"GUIAS USADAS: {', '.join(guias_selecionadas)}")

    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho_saida)
    return caminho_saida

# =========================
# ENCAMINHAMENTOS
# =========================
def quebrar_complementares(texto):
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return []
    texto = str(texto).strip()
    if not texto or texto.lower() == "nan":
        return []
    # A base pode vir separada por ;, quebra de linha, vírgula ou barra vertical.
    partes = re.split(r"[;\n\r|]+", texto)
    if len(partes) == 1 and "," in texto and not re.search(r"\d+,\d+", texto):
        partes = texto.split(",")
    return [x.strip().upper() for x in partes if str(x).strip()]


def _cnpj_pasta_encaminhamento(valor):
    """Nome seguro da pasta/zip da empresa: somente números do CNPJ."""
    numeros = somente_numeros(valor)
    return numeros if numeros else "SEM_CNPJ"


def _valor_linha_encaminhamento(row, coluna):
    if not coluna:
        return ""
    try:
        valor = row[coluna]
    except Exception:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass
    return str(valor).strip()


def _nome_arquivo_unico(pasta, nome_base, extensao):
    pasta = Path(pasta)
    nome_limpo = limpar_nome_pasta_arquivo(nome_base or "SEM NOME") or "SEM NOME"
    destino = pasta / f"{nome_limpo}.{extensao}"
    contador = 1
    while destino.exists():
        destino = pasta / f"{nome_limpo} ({contador}).{extensao}"
        contador += 1
    return destino


def _data_geracao_encaminhamento():
    """Data local de Belém/Macapá no formato exibido no encaminhamento."""
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("America/Belem")).strftime("%d/%m/%Y")
    except Exception:
        return datetime.now().strftime("%d/%m/%Y")


def _gerar_encaminhamento_docx(contexto, destino):
    template = DocxTemplate(TEMPLATE_PATH)
    template.render(contexto)
    template.save(str(destino))


def _soffice_executavel():
    """Localiza o LibreOffice usado para converter o mesmo DOCX em PDF."""
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        raise RuntimeError(
            "LibreOffice não está instalado no servidor. "
            "Ele é necessário para gerar o PDF com exatamente o mesmo modelo do Word."
        )
    return exe


def _converter_docx_em_lote_para_pdf(caminhos_docx, pasta_destino, tamanho_lote=40):
    """Converte DOCX já renderizados do modelo oficial para PDF.

    O PDF não é redesenhado pelo sistema. Primeiro geramos o mesmo arquivo Word
    utilizado na opção DOCX e depois o LibreOffice apenas o exporta para PDF.
    Dessa forma Word e PDF compartilham layout, logotipo, tabelas, espaçamentos,
    endereço, data e demais elementos do mesmo template.
    """
    caminhos = [Path(c) for c in caminhos_docx]
    if not caminhos:
        return []

    destino = Path(pasta_destino)
    destino.mkdir(parents=True, exist_ok=True)
    soffice = _soffice_executavel()
    gerados = []

    for inicio in range(0, len(caminhos), max(1, int(tamanho_lote))):
        lote = caminhos[inicio:inicio + max(1, int(tamanho_lote))]
        perfil_dir = Path(tempfile.mkdtemp(prefix="edge_lo_profile_"))
        try:
            cmd = [
                soffice,
                f"-env:UserInstallation={perfil_dir.resolve().as_uri()}",
                "--headless",
                "--nologo",
                "--nofirststartwizard",
                "--norestore",
                "--convert-to", "pdf",
                "--outdir", str(destino),
                *[str(c.resolve()) for c in lote],
            ]
            proc = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=max(180, 15 * len(lote)),
            )
            if proc.returncode != 0:
                detalhe = (proc.stderr or proc.stdout or "erro desconhecido").strip()
                raise RuntimeError(f"Falha ao converter os encaminhamentos para PDF: {detalhe}")

            faltantes = []
            for caminho_docx in lote:
                pdf_esperado = destino / f"{caminho_docx.stem}.pdf"
                if not pdf_esperado.exists() or pdf_esperado.stat().st_size == 0:
                    faltantes.append(caminho_docx.name)
                else:
                    gerados.append(pdf_esperado)
            if faltantes:
                detalhe = (proc.stderr or proc.stdout or "").strip()
                raise RuntimeError(
                    "O LibreOffice não gerou todos os PDFs esperados. "
                    f"Arquivos: {', '.join(faltantes[:5])}. {detalhe}"
                )
        finally:
            shutil.rmtree(perfil_dir, ignore_errors=True)

    return gerados


def _gerar_encaminhamento_pdf(contexto, destino):
    """Compatibilidade: gera o Word oficial e exporta esse mesmo documento para PDF."""
    destino = Path(destino)
    with tempfile.TemporaryDirectory(prefix="edge_enc_pdf_") as tmp:
        docx_temp = Path(tmp) / f"{destino.stem}.docx"
        _gerar_encaminhamento_docx(contexto, docx_temp)
        gerados = _converter_docx_em_lote_para_pdf([docx_temp], destino.parent, tamanho_lote=1)
        pdf_gerado = gerados[0]
        if pdf_gerado.resolve() != destino.resolve():
            if destino.exists():
                destino.unlink()
            shutil.move(str(pdf_gerado), str(destino))

def gerar_encaminhamentos(file, formato_saida="docx"):
    file.seek(0)
    df = pd.read_excel(file)

    col_empresa = encontrar_coluna(df, ["empresa"], obrigatoria=True)
    col_cnpj = encontrar_coluna(df, ["cnpj", "cpf"], obrigatoria=True)
    col_nome = encontrar_coluna(df, ["funcionario", "funcionário", "nome"], obrigatoria=True)
    col_funcao = encontrar_coluna(df, ["funcao", "função", "cargo"])
    col_comp = encontrar_coluna(df, ["complementares", "exames", "exames_obg", "exames obrigatorios", "exames obrigatórios"], obrigatoria=False)

    formato_saida = str(formato_saida or "docx").strip().lower()
    if formato_saida not in {"docx", "pdf"}:
        formato_saida = "docx"

    temp_dir = tempfile.mkdtemp()
    empresas_root = Path(temp_dir) / "empresas"
    empresas_root.mkdir(parents=True, exist_ok=True)

    registros_por_empresa = {}
    relatorio = []
    data_geracao = _data_geracao_encaminhamento()

    for _, row in df.iterrows():
        empresa = _valor_linha_encaminhamento(row, col_empresa)
        funcionario = _valor_linha_encaminhamento(row, col_nome)
        cnpj_raw = _valor_linha_encaminhamento(row, col_cnpj)
        funcao = _valor_linha_encaminhamento(row, col_funcao)
        complementares_txt = _valor_linha_encaminhamento(row, col_comp)
        if not funcionario:
            continue

        cnpj_pasta = _cnpj_pasta_encaminhamento(cnpj_raw)
        cnpj_formatado = formatar_documento(cnpj_raw) or cnpj_raw
        registros_por_empresa.setdefault(cnpj_pasta, []).append({
            "empresa": empresa.upper(),
            "cnpj": cnpj_formatado,
            "funcionario": funcionario.upper(),
            "funcao": funcao.upper(),
            "complementares": quebrar_complementares(complementares_txt),
        })

    if not registros_por_empresa:
        raise ValueError("Nenhum encaminhamento foi encontrado na planilha. Confira as colunas EMPRESA, CNPJ, NOME e COMPLEMENTARES.")

    for cnpj_pasta, registros in registros_por_empresa.items():
        pasta_empresa = empresas_root / cnpj_pasta
        pasta_empresa.mkdir(parents=True, exist_ok=True)
        docx_para_converter = []
        temp_docx_dir = None
        if formato_saida == "pdf":
            temp_docx_dir = Path(temp_dir) / "docx_para_pdf" / cnpj_pasta
            temp_docx_dir.mkdir(parents=True, exist_ok=True)

        for item in registros:
            comps = {f"comp{i+1}": item["complementares"][i] if i < len(item["complementares"]) else "" for i in range(9)}
            contexto = {
                "empresa": item["empresa"],
                "cnpj": item["cnpj"],
                "funcionario": item["funcionario"],
                "funcao": item["funcao"],
                "data_geracao": data_geracao,
                **comps,
            }
            base_nome = f"ENCAMINHAMENTO {contexto['funcionario'] or 'SEM NOME'}"
            if formato_saida == "pdf":
                # O nome é definido pela saída PDF; o DOCX temporário usa o mesmo nome-base.
                destino_pdf = _nome_arquivo_unico(pasta_empresa, base_nome, "pdf")
                docx_temp = temp_docx_dir / f"{destino_pdf.stem}.docx"
                _gerar_encaminhamento_docx(contexto, docx_temp)
                docx_para_converter.append(docx_temp)
            else:
                destino = _nome_arquivo_unico(pasta_empresa, base_nome, "docx")
                _gerar_encaminhamento_docx(contexto, destino)

        if formato_saida == "pdf":
            _converter_docx_em_lote_para_pdf(docx_para_converter, pasta_empresa)
            shutil.rmtree(temp_docx_dir, ignore_errors=True)

        relatorio.append(f"{cnpj_pasta}: {len(registros)} encaminhamento(s)")

    # ZIP principal: dentro dele vai 1 ZIP por empresa/CNPJ, e dentro de cada ZIP fica a pasta do CNPJ com os encaminhamentos.
    zip_principal = Path(temp_dir) / "encaminhamentos.zip"
    with zipfile.ZipFile(zip_principal, "w", zipfile.ZIP_DEFLATED) as zip_out:
        for pasta_empresa in sorted(empresas_root.iterdir(), key=lambda p: p.name):
            if not pasta_empresa.is_dir():
                continue
            zip_empresa_path = Path(temp_dir) / f"{pasta_empresa.name}.zip"
            with zipfile.ZipFile(zip_empresa_path, "w", zipfile.ZIP_DEFLATED) as zip_empresa:
                for arquivo in sorted(pasta_empresa.iterdir(), key=lambda p: p.name):
                    if arquivo.is_file():
                        zip_empresa.write(arquivo, f"{pasta_empresa.name}/{arquivo.name}")
            zip_out.write(zip_empresa_path, zip_empresa_path.name)
        resumo_texto = (
            "ENCAMINHAMENTOS GERADOS\n"
            "========================\n"
            f"Formato dos arquivos: {formato_saida.upper()}\n"
            f"Empresas/CNPJs: {len(registros_por_empresa)}\n"
            f"Total de encaminhamentos: {sum(len(v) for v in registros_por_empresa.values())}\n\n"
            + "\n".join(relatorio)
        )
        zip_out.writestr("RESUMO_ENCAMINHAMENTOS.txt", resumo_texto)
    return str(zip_principal)

# =========================
# RENUMERADOR
# =========================
def allowed_renum_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in RENUM_ALLOWED_EXTENSIONS

def _docx_xml_tree(caminho_entrada: str):
    import xml.etree.ElementTree as ET
    ns_map = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'v': 'urn:schemas-microsoft-com:vml',
        'o': 'urn:schemas-microsoft-com:office:office',
        'w10': 'urn:schemas-microsoft-com:office:word',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    }
    for prefix, uri in ns_map.items():
        ET.register_namespace(prefix, uri)
    with zipfile.ZipFile(caminho_entrada, 'r') as zin:
        xml_bytes = zin.read('word/document.xml')
    return ET, ET.fromstring(xml_bytes), ns_map


def _docx_xml_paragraphs(root, ns_map):
    w = '{%s}' % ns_map['w']
    return list(root.iter(w + 'p'))


def _paragraph_text_xml_et(p, ns_map) -> str:
    w = '{%s}' % ns_map['w']
    return ''.join((t.text or '') for t in p.iter(w + 't'))


def _paragraph_text_clean_et(p, ns_map) -> str:
    return re.sub(r"\s+", " ", _paragraph_text_xml_et(p, ns_map)).strip()


def _ensure_run_property_et(run, ns_map):
    w = '{%s}' % ns_map['w']
    rpr = run.find(w + 'rPr')
    if rpr is None:
        rpr = ET.Element(w + 'rPr')
        run.insert(0, rpr)
    return rpr


def _set_run_bold_et(ET, run, ns_map, enabled: bool):
    w = '{%s}' % ns_map['w']
    rpr = run.find(w + 'rPr')
    if rpr is None:
        rpr = ET.Element(w + 'rPr')
        run.insert(0, rpr)
    b = rpr.find(w + 'b')
    if b is None:
        b = ET.Element(w + 'b')
        rpr.append(b)
    val_attr = w + 'val'
    if enabled:
        b.attrib.pop(val_attr, None)
    else:
        b.set(val_attr, '0')


def _set_paragraph_bold_et(ET, p, ns_map, enabled: bool):
    w = '{%s}' % ns_map['w']
    for run in p.iter(w + 'r'):
        _set_run_bold_et(ET, run, ns_map, enabled)


def _replace_paragraph_text_et(ET, p, ns_map, novo_texto: str, bold: bool | None = None):
    w = '{%s}' % ns_map['w']
    texts = list(p.iter(w + 't'))
    if not texts:
        return False
    texts[0].text = novo_texto
    for t in texts[1:]:
        t.text = ''
    if bold is not None:
        _set_paragraph_bold_et(ET, p, ns_map, bold)
    return True


def _looks_like_nota_balcao(texto: str) -> bool:
    return normalize_text(texto) == 'NOTA DE BALCAO'


def _is_receipt_number(texto: str) -> bool:
    return bool(re.fullmatch(r"\d{1,8}", texto.strip()))


def _write_docx_preserving_package(caminho_entrada: str, caminho_saida: str, document_xml: bytes):
    Path(caminho_saida).parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(caminho_entrada, 'r') as zin:
        with zipfile.ZipFile(caminho_saida, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item, document_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))


def _paragraph_text_nodes_with_offsets(p, ns_map):
    """Retorna o texto completo do parágrafo e o mapa de offsets por nó <w:t>.

    Isso permite trocar apenas um trecho específico do texto sem reconstruir o
    parágrafo inteiro. É essencial para DOCX reais, onde o Word divide um único
    número em vários runs, por exemplo: "1" + "4" + "2".
    """
    w = '{%s}' % ns_map['w']
    text_nodes = list(p.iter(w + 't'))
    pieces = []
    offsets = []
    cursor = 0
    for node in text_nodes:
        value = node.text or ''
        pieces.append(value)
        offsets.append((node, cursor, cursor + len(value)))
        cursor += len(value)
    return ''.join(pieces), offsets


def _replace_text_range_in_paragraph(p, ns_map, start: int, end: int, replacement: str, bold: bool | None = None) -> bool:
    """Substitui um intervalo de caracteres preservando o máximo da formatação.

    O replacement é colocado no primeiro nó de texto atingido e o trecho antigo
    é apagado dos demais nós. Assim evitamos o erro clássico de usar
    paragraph.text = ..., que destrói a formatação/tabelas do Word.
    """
    full_text, offsets = _paragraph_text_nodes_with_offsets(p, ns_map)
    if start < 0 or end <= start or end > len(full_text):
        return False

    first_written = False
    for node, node_start, node_end in offsets:
        if node_end <= start or node_start >= end:
            continue
        local_start = max(0, start - node_start)
        local_end = min(node_end - node_start, end - node_start)
        current = node.text or ''
        before = current[:local_start]
        after = current[local_end:]
        if not first_written:
            node.text = before + replacement + after
            first_written = True
        else:
            node.text = before + after

    if bold is not None:
        import xml.etree.ElementTree as _ET
        _set_paragraph_bold_et(_ET, p, ns_map, bold)
    return first_written


def _receipt_number_match_in_paragraph(texto: str):
    """Encontra o número do recibo dentro de um parágrafo.

    Casos aceitos:
    1. Parágrafo só com número: "00142".
    2. Número grudado ao total: "TOTAL R$: 150,00142".

    O segundo caso apareceu nos recibos reais e era a causa de arquivos não
    renumerados. A regra só roda quando o parágrafo possui TOTAL/R$, reduzindo
    risco de alterar CNPJ, datas ou valores indevidos.
    """
    raw = texto or ''
    stripped = raw.strip()
    if _is_receipt_number(stripped):
        start = raw.find(stripped)
        return start, start + len(stripped), stripped, 'numero_isolado'

    normalized = normalize_text(raw)
    if 'TOTAL' in normalized and ('R$' in raw.upper() or 'R$:' in raw.upper() or 'R$' in normalized):
        # Ex.: TOTAL R$: 150,00TOTAL R$: 150,00142
        match = re.search(r"(?:R\$\s*:?\s*)?\d{1,3}(?:\.\d{3})*,\d{2}(\d{1,8})\s*$", raw, re.IGNORECASE)
        if match:
            number = match.group(1)
            return match.start(1), match.end(1), number, 'numero_grudado_ao_total'

    return None


def _find_receipt_number_after_nota(paragraphs, ns_map, nota_index: int, lookahead: int = 12):
    """Busca o número do recibo depois de NOTA DE BALCÃO.

    A busca para ao encontrar Data/Cliente ou outra NOTA, mas também tolera
    documentos onde o número aparece grudado no total antes da data.
    """
    stop_re = re.compile(r"^\s*(Data:|CLIENTE:)", re.IGNORECASE)
    best = None
    for j in range(nota_index + 1, min(nota_index + lookahead + 1, len(paragraphs))):
        texto = _paragraph_text_clean_et(paragraphs[j], ns_map)
        if _looks_like_nota_balcao(texto):
            break
        match = _receipt_number_match_in_paragraph(texto)
        if match:
            return j, match
        if stop_re.match(texto) and best is None:
            break
    return best


def encontrar_numeros_recibos_xml(paragraphs, ns_map):
    encontrados = []
    pendentes = []
    for i, p in enumerate(paragraphs):
        if _looks_like_nota_balcao(_paragraph_text_clean_et(p, ns_map)):
            found = _find_receipt_number_after_nota(paragraphs, ns_map, i)
            if found:
                paragraph_index, match = found
                start, end, numero, tipo = match
                encontrados.append({
                    'nota_index': i,
                    'paragraph_index': paragraph_index,
                    'start': start,
                    'end': end,
                    'numero': numero,
                    'tipo': tipo,
                })
            else:
                pendentes.append(i)
    return encontrados, pendentes


def encontrar_ultimo_numero_xml(paragraphs, ns_map) -> str:
    encontrados, _pendentes = encontrar_numeros_recibos_xml(paragraphs, ns_map)
    numeros = [item['numero'] for item in encontrados if _is_receipt_number(item['numero'])]
    if numeros:
        return max(numeros, key=lambda x: int(x))
    return "0"


def renumerar_documento(caminho_entrada: str, caminho_saida: str, nova_data: str):
    ET, root, ns_map = _docx_xml_tree(caminho_entrada)
    paragraphs = _docx_xml_paragraphs(root, ns_map)
    encontrados, pendentes = encontrar_numeros_recibos_xml(paragraphs, ns_map)
    numeros_originais = [item['numero'] for item in encontrados]
    ultimo = max(numeros_originais, key=lambda x: int(x)) if numeros_originais else "0"
    tamanho = max([len(n) for n in numeros_originais] + [len(ultimo), 1])
    numero_atual = int(ultimo) + 1
    alterados = 0
    datas_alteradas = 0
    avisos = []

    if pendentes:
        avisos.append(f"{len(pendentes)} NOTA(S) DE BALCÃO sem número detectado")

    # Remove negrito do título NOTA DE BALCÃO, mantendo o padrão usado nos recibos.
    for item in encontrados:
        _set_paragraph_bold_et(ET, paragraphs[item['nota_index']], ns_map, False)

    for item in encontrados:
        p = paragraphs[item['paragraph_index']]
        novo_num = str(numero_atual).zfill(max(tamanho, len(item['numero'])))
        # Recalcula o match no momento da substituição porque offsets podem mudar
        # quando há mais de uma troca no mesmo parágrafo.
        texto_atual, _offsets = _paragraph_text_nodes_with_offsets(p, ns_map)
        match = _receipt_number_match_in_paragraph(texto_atual)
        if not match:
            avisos.append(f"Número original {item['numero']} não pôde ser reencontrado no parágrafo")
            continue
        start, end, _numero_detectado, _tipo = match
        if _replace_text_range_in_paragraph(p, ns_map, start, end, novo_num, bold=True):
            numero_atual += 1
            alterados += 1
        else:
            avisos.append(f"Falha ao substituir número {item['numero']}")

    data_padrao = re.compile(r"^\s*Data:\s*\d{2}/\d{2}/\d{4}\s*$", re.IGNORECASE)
    for p in paragraphs:
        texto = _paragraph_text_clean_et(p, ns_map)
        if data_padrao.match(texto):
            if _replace_paragraph_text_et(ET, p, ns_map, f"Data: {nova_data}", bold=True):
                datas_alteradas += 1

    if encontrados and datas_alteradas != alterados:
        avisos.append(f"Quantidade de datas alteradas ({datas_alteradas}) diferente de recibos renumerados ({alterados})")

    xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    _write_docx_preserving_package(caminho_entrada, caminho_saida, xml_bytes)
    return alterados, ultimo, datas_alteradas, avisos

# =========================
# RECIBO eSOCIAL
# =========================
def normalize_company_name(value) -> str:
    text = normalize_text(value)
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\b(LTDA|EIRELI|ME|EPP|S A|SA|S/S|SS|MATRIZ|FILIAL)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ESOCIAL_ALLOWED_EXTENSIONS


def extract_cnpj(text: str) -> str:
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    if isinstance(text, (int, float)) and not isinstance(text, bool):
        try:
            text = str(int(text))
        except Exception:
            text = str(text)
    digits = re.sub(r"\D", "", str(text))
    return digits[:14] if len(digits) >= 14 else ""


def format_cnpj(cnpj: str) -> str:
    digits = re.sub(r"\D", "", cnpj or "")
    if len(digits) != 14:
        return "CNPJ NÃO INFORMADO"
    return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"


def format_cnpj_filename(cnpj: str) -> str:
    """Formato usado na nomenclatura histórica dos recibos: 00.000.000.0000-00."""
    digits = re.sub(r"\D", "", cnpj or "")
    if len(digits) != 14:
        return "CNPJ NÃO INFORMADO"
    return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}.{digits[8:12]}-{digits[12:]}"


def list_sheets(path: str):
    try:
        xl = pd.ExcelFile(path)
        return xl.sheet_names or ["Planilha principal"]
    except Exception as exc:
        raise RuntimeError(f"Não foi possível ler as guias de {os.path.basename(path)}. Erro: {exc}") from exc


def _find_column_optional(df: pd.DataFrame, expected_names: list[str]) -> str | None:
    normalized = {normalize_text(col): col for col in df.columns}
    for name in expected_names:
        norm = normalize_text(name)
        if norm in normalized:
            return normalized[norm]
    return None


def find_column(df: pd.DataFrame, expected_names: list[str]) -> str:
    found = _find_column_optional(df, expected_names)
    if found:
        return found
    raise KeyError(
        f"Coluna não encontrada. Esperado um destes nomes: {expected_names}. "
        f"Colunas encontradas: {list(df.columns)}"
    )


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy().dropna(axis=1, how="all").dropna(axis=0, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _normalize_marker(value) -> str:
    text = normalize_text(value)
    return re.sub(r"[^A-Z0-9]+", " ", text).strip()


def _row_has_ok_esocial(row: pd.Series) -> bool:
    # Aceita OK E-SOCIAL, OK E SOCIAL e complementos como OK E-SOCIAL/P- ANTIGA.
    # Também reconhece o legado E-SOCIAL OK sem prejudicar o fluxo atual.
    for value in row.tolist():
        marker = _normalize_marker(value)
        if "OK E SOCIAL" in marker or "E SOCIAL OK" in marker:
            return True
    return False


def _normalize_person_name(value) -> str:
    text = normalize_text(value)
    text = re.sub(r"[^A-Z0-9 ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _safe_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _cpf_text(value) -> str:
    text = _safe_text(value)
    digits = re.sub(r"\D", "", text)
    if digits and len(digits) <= 11:
        return digits.zfill(11)
    return text


def _parse_excel_date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            number = float(value)
            if 20000 <= number <= 80000:
                return (datetime(1899, 12, 30) + timedelta(days=number)).date()
        except Exception:
            pass
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text[:10], fmt).date()
        except Exception:
            pass
    try:
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
        if not pd.isna(parsed):
            return parsed.date()
    except Exception:
        pass
    return None


def _format_date(value) -> str:
    parsed = _parse_excel_date(value)
    return parsed.strftime("%d/%m/%Y") if parsed else _safe_text(value)


def _sheet_month_year(sheet_name: str) -> tuple[str, int | None]:
    normalized = normalize_text(sheet_name)
    month = ""
    for key, label in ESOCIAL_MONTHS.items():
        if re.search(rf"(^|[^A-Z]){re.escape(key)}([^A-Z]|$)", normalized):
            month = label
            break
    year_match = re.search(r"(20\d{2})", str(sheet_name))
    year = int(year_match.group(1)) if year_match else None
    return month, year


def _strip_cnpj_from_company(company_text: str, cnpj: str) -> str:
    company = _safe_text(company_text)
    digits = re.sub(r"\D", "", cnpj or "")
    if digits:
        flexible = r"\D*".join(re.escape(ch) for ch in digits)
        company = re.sub(flexible, "", company, count=1)
    company = re.sub(r"\s*[-–—|/]\s*$", "", company).strip(" -–—|/")
    return company or "EMPRESA"


def read_esocial_base_rows(base_file: str, base_sheet: str) -> pd.DataFrame:
    if not base_sheet:
        raise ValueError("Selecione a guia/mês da planilha base.")
    try:
        df = pd.read_excel(base_file, sheet_name=base_sheet, dtype=object)
    except Exception as exc:
        raise RuntimeError(f"Não foi possível ler a guia '{base_sheet}' da planilha base. Erro: {exc}") from exc
    df = prepare_dataframe(df)
    if df.empty:
        raise ValueError("A guia selecionada da planilha base está vazia.")

    employee_col = find_column(df, ["FUNCIONÁRIO", "FUNCIONARIO", "NOME", "COLABORADOR"])
    company_col = _find_column_optional(df, ["SETOR", "EMPRESA", "UNIDADE", "RAZÃO SOCIAL", "RAZAO SOCIAL"])
    cnpj_col = _find_column_optional(df, ["CNPJ", "CNPJ EMPRESA", "CNPJ DA EMPRESA"])
    date_col = _find_column_optional(df, ["DATA", "DATA DO EXAME", "DATA EXAME"])
    if not company_col and not cnpj_col:
        raise KeyError("Não foi possível identificar a empresa/CNPJ na planilha base.")

    eligible = df[df.apply(_row_has_ok_esocial, axis=1)].copy()
    if eligible.empty:
        raise ValueError("Não encontrei nenhuma linha com 'OK E-SOCIAL' na guia selecionada.")

    rows = []
    for idx, row in eligible.iterrows():
        company_text = _safe_text(row.get(company_col, "")) if company_col else ""
        cnpj = extract_cnpj(row.get(cnpj_col, "")) if cnpj_col else ""
        if not cnpj:
            cnpj = extract_cnpj(company_text)
        employee = _safe_text(row.get(employee_col, ""))
        if len(cnpj) != 14 or not employee:
            continue
        rows.append({
            "BASE_ROW": int(idx) + 2 if isinstance(idx, int) else len(rows) + 2,
            "CNPJ": cnpj,
            "EMPRESA_BASE": company_text,
            "EMPRESA_NOME": _strip_cnpj_from_company(company_text, cnpj),
            "FUNCIONARIO_BASE": employee,
            "NOME_KEY": _normalize_person_name(employee),
            "BASE_DATE": _parse_excel_date(row.get(date_col, "")) if date_col else None,
        })
    result = pd.DataFrame(rows)
    if result.empty:
        raise ValueError("As linhas com OK E-SOCIAL não possuem CNPJ de empresa e funcionário válidos.")
    return result


def _score_esocial_export_sheet(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return -1
    cols = {normalize_text(c) for c in df.columns}
    required_groups = [
        {"CNPJ", "CNPJ EMPRESA", "CNPJ DA EMPRESA"},
        {"FUNCIONARIO", "FUNCIONÁRIO", "NOME", "COLABORADOR"},
        {"RECIBO", "NUMERO DO RECIBO", "NÚMERO DO RECIBO", "RECIBO ESOCIAL", "RECIBO E-SOCIAL"},
    ]
    score = 0
    for group in required_groups:
        if cols.intersection(group):
            score += 100
    for wanted in ["EVENTO", "EMPRESA", "CPF", "MATRICULA", "MATRÍCULA", "DATA REF.", "DATA REF", "STATUS", "DATA ENVIO"]:
        if wanted in cols:
            score += 10
    score += min(len(df), 500)
    return score


def read_esocial_export_file(path: str) -> tuple[pd.DataFrame, str]:
    try:
        xl = pd.ExcelFile(path)
    except Exception as exc:
        raise RuntimeError(f"Não foi possível abrir {os.path.basename(path)}. Erro: {exc}") from exc

    best_df = None
    best_sheet = ""
    best_score = -1
    for sheet in xl.sheet_names:
        try:
            candidate = pd.read_excel(path, sheet_name=sheet, dtype=object)
            candidate = prepare_dataframe(candidate)
        except Exception:
            continue
        score = _score_esocial_export_sheet(candidate)
        if score > best_score:
            best_score = score
            best_df = candidate
            best_sheet = sheet

    if best_df is None or best_score < 300:
        raise ValueError(
            f"{os.path.basename(path)} não possui uma guia de envios com CNPJ, Funcionário e Recibo."
        )

    df = best_df
    evento_col = _find_column_optional(df, ["EVENTO", "TIPO EVENTO"])
    empresa_col = _find_column_optional(df, ["EMPRESA", "RAZÃO SOCIAL", "RAZAO SOCIAL"])
    cnpj_col = find_column(df, ["CNPJ", "CNPJ EMPRESA", "CNPJ DA EMPRESA"])
    funcionario_col = find_column(df, ["FUNCIONÁRIO", "FUNCIONARIO", "NOME", "COLABORADOR"])
    cpf_col = _find_column_optional(df, ["CPF"])
    matricula_col = _find_column_optional(df, ["MATRÍCULA", "MATRICULA"])
    data_ref_col = _find_column_optional(df, ["DATA REF.", "DATA REF", "DATA REFERÊNCIA", "DATA REFERENCIA", "DATA"])
    status_col = _find_column_optional(df, ["STATUS"])
    data_envio_col = _find_column_optional(df, ["DATA ENVIO", "DATA DE ENVIO"])
    recibo_col = find_column(df, ["RECIBO", "NÚMERO DO RECIBO", "NUMERO DO RECIBO", "RECIBO ESOCIAL", "RECIBO E-SOCIAL"])

    rows = []
    for _, row in df.iterrows():
        cnpj = extract_cnpj(row.get(cnpj_col, ""))
        employee = _safe_text(row.get(funcionario_col, ""))
        if len(cnpj) != 14 or not employee:
            continue
        evento = _safe_text(row.get(evento_col, "")) if evento_col else ""
        rows.append({
            "EVENTO": evento,
            "EMPRESA": _safe_text(row.get(empresa_col, "")) if empresa_col else "",
            "CNPJ": cnpj,
            "FUNCIONARIO": employee,
            "NOME_KEY": _normalize_person_name(employee),
            "CPF": _cpf_text(row.get(cpf_col, "")) if cpf_col else "",
            "MATRICULA": _safe_text(row.get(matricula_col, "")) if matricula_col else "",
            "DATA_REF_DATE": _parse_excel_date(row.get(data_ref_col, "")) if data_ref_col else None,
            "DATA_REF": _format_date(row.get(data_ref_col, "")) if data_ref_col else "",
            "STATUS": _safe_text(row.get(status_col, "")) if status_col else "",
            "DATA_ENVIO_DATE": _parse_excel_date(row.get(data_envio_col, "")) if data_envio_col else None,
            "DATA_ENVIO": _format_date(row.get(data_envio_col, "")) if data_envio_col else "",
            "RECIBO": _safe_text(row.get(recibo_col, "")),
            "ARQUIVO_ORIGEM": os.path.basename(path),
            "GUIA_ORIGEM": best_sheet,
        })

    result = pd.DataFrame(rows)
    if result.empty:
        raise ValueError(f"{os.path.basename(path)} não possui registros de funcionários válidos.")

    # A base clínica representa ASO/S-2220. Se a exportação contiver S-2220, ignora outros eventos.
    if "EVENTO" in result.columns:
        event_key = result["EVENTO"].map(lambda x: re.sub(r"[^A-Z0-9]", "", normalize_text(x)))
        if (event_key == "S2220").any():
            result = result[event_key == "S2220"].copy()
    return result.reset_index(drop=True), best_sheet


def _candidate_priority(row: pd.Series, base_date) -> tuple:
    ref_date = row.get("DATA_REF_DATE")
    exact = int(bool(base_date and ref_date and ref_date == base_date))
    same_month = int(bool(base_date and ref_date and ref_date.year == base_date.year and ref_date.month == base_date.month))
    authorized = int("AUTORIZ" in normalize_text(row.get("STATUS", "")))
    has_receipt = int(bool(_safe_text(row.get("RECIBO", ""))))
    distance = 999999
    if base_date and ref_date:
        try:
            distance = abs((ref_date - base_date).days)
        except Exception:
            pass
    ref_ord = ref_date.toordinal() if ref_date else 0
    send_date = row.get("DATA_ENVIO_DATE")
    send_ord = send_date.toordinal() if send_date else 0
    return (exact, same_month, authorized, has_receipt, -distance, ref_ord, send_ord)


def select_esocial_rows_for_company(base_company: pd.DataFrame, export_company: pd.DataFrame):
    selected = []
    missing = []
    used_indexes = set()

    for _, base_row in base_company.sort_values(["BASE_DATE", "BASE_ROW"], na_position="last").iterrows():
        candidates = export_company[export_company["NOME_KEY"] == base_row["NOME_KEY"]]
        if candidates.empty:
            missing.append(base_row["FUNCIONARIO_BASE"])
            continue

        unused = candidates[~candidates.index.isin(used_indexes)]
        pool = unused if not unused.empty else candidates
        ranked = sorted(
            [(idx, row) for idx, row in pool.iterrows()],
            key=lambda item: _candidate_priority(item[1], base_row.get("BASE_DATE")),
            reverse=True,
        )
        chosen_idx, chosen = ranked[0]
        used_indexes.add(chosen_idx)
        selected.append(chosen.to_dict())

    if not selected:
        return pd.DataFrame(columns=export_company.columns), missing

    result = pd.DataFrame(selected)
    dedupe_cols = [c for c in ["CNPJ", "NOME_KEY", "DATA_REF", "RECIBO", "EVENTO"] if c in result.columns]
    if dedupe_cols:
        result = result.drop_duplicates(subset=dedupe_cols, keep="first")
    return result.reset_index(drop=True), missing


def make_paragraph(text: str, style: ParagraphStyle) -> Paragraph:
    text = _safe_text(text)
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
    return Paragraph(text, style)


def build_esocial_receipt_pdf(df: pd.DataFrame, pdf_path: str, company_name: str, company_cnpj: str):
    if df.empty:
        raise ValueError("Não há funcionários em comum para gerar o recibo da empresa.")

    page_width, _ = landscape(A4)
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(A4),
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=18 * mm,
        bottomMargin=14 * mm,
    )
    styles = getSampleStyleSheet()
    body_style = ParagraphStyle(
        "EsocialCell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=7.6,
        leading=8.8,
        alignment=0,
        spaceBefore=0,
        spaceAfter=0,
        textColor=colors.black,
    )
    header_style = ParagraphStyle(
        "EsocialHeader",
        parent=body_style,
        fontName="Helvetica-Bold",
        textColor=colors.white,
        fontSize=7.7,
        leading=8.8,
    )

    headers = ["Evento", "Empresa", "Funcionário", "CPF", "Matrícula", "Data Ref.", "Status", "Data Envio", "Recibo"]
    full_company = f"{company_name} ({format_cnpj(company_cnpj)})"
    table_data = [[make_paragraph(h, header_style) for h in headers]]
    for _, row in df.iterrows():
        # Mantém o mesmo padrão visual do recibo de referência: Razão Social (CNPJ).
        company_value = full_company
        table_data.append([
            make_paragraph(row.get("EVENTO") or "S-2220", body_style),
            make_paragraph(company_value, body_style),
            make_paragraph(row.get("FUNCIONARIO"), body_style),
            make_paragraph(row.get("CPF"), body_style),
            make_paragraph(row.get("MATRICULA"), body_style),
            make_paragraph(row.get("DATA_REF"), body_style),
            make_paragraph(row.get("STATUS"), body_style),
            make_paragraph(row.get("DATA_ENVIO"), body_style),
            make_paragraph(row.get("RECIBO"), body_style),
        ])

    usable = page_width - doc.leftMargin - doc.rightMargin
    weights = [5, 30, 19, 9, 9, 7, 7, 7, 14]
    total = sum(weights)
    widths = [usable * value / total for value in weights]
    table = Table(table_data, colWidths=widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#23496D")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.65, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.3),
        ("TOPPADDING", (0, 0), (-1, -1), 2.3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.3),
    ]))
    doc.build([table])


def build_pdf(df: pd.DataFrame, pdf_path: str, title: str):
    """Compatibilidade com o serviço antigo; usa o novo layout quando chamado diretamente."""
    cnpj = extract_cnpj(title)
    company = _strip_cnpj_from_company(title, cnpj) if cnpj else (_safe_text(title) or "EMPRESA")
    build_esocial_receipt_pdf(df, pdf_path, company, cnpj)


def build_esocial_pdf_filename(company_name: str, company_cnpj: str, pdf_month: str) -> str:
    return sanitize_filename(
        f"{company_name} - {format_cnpj_filename(company_cnpj)} - {pdf_month}"
    ) + ".pdf"


def create_output_folder(base_output_dir: str) -> str:
    folder_path = os.path.join(base_output_dir, "RECIBOS ESOCIAL")
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def create_structure(base_folder: str):
    # Mantido para compatibilidade com o módulo de serviços.
    pdf_folder = os.path.join(base_folder, "PDFs")
    log_folder = os.path.join(base_folder, "Logs")
    os.makedirs(pdf_folder, exist_ok=True)
    os.makedirs(log_folder, exist_ok=True)
    return pdf_folder, log_folder


def create_zip_from_folder(folder: str) -> str:
    zip_path = unique_path(folder.rstrip("/\\") + ".zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(folder):
            for filename in files:
                if filename.lower().endswith(".zip"):
                    continue
                full = os.path.join(root, filename)
                zf.write(full, os.path.relpath(full, folder))
    return zip_path


def create_esocial_zip(folder: str, month: str, year: int | None = None) -> str:
    label = f"RECIBOS ESOCIAL - {month}" + (f" {year}" if year else "")
    zip_path = unique_path(os.path.join(os.path.dirname(folder), sanitize_filename(label) + ".zip"))
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename in sorted(os.listdir(folder)):
            full = os.path.join(folder, filename)
            if os.path.isfile(full) and not filename.lower().endswith(".zip"):
                zf.write(full, filename)
    return zip_path


def export_summary_excel(rows: list[dict], excel_path: str):
    """Compatibilidade com versões anteriores; não é usado no novo Recibo eSocial."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"
    headers = ["EMPRESA", "CNPJ", "STATUS", "TOTAL BASE EMPRESA", "TOTAL ENCONTRADO", "MOTIVO", "PDF GERADO"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append([
            row.get("empresa", ""), row.get("cnpj", ""), row.get("status", ""),
            row.get("total_base", 0), row.get("total_encontrado", 0),
            row.get("motivo", ""), row.get("pdf", ""),
        ])
    wb.save(excel_path)


def process_esocial_receipts(base_file: str, base_sheet: str, export_files: list[str], output_folder: str, progress=None):
    month, year = _sheet_month_year(base_sheet)
    if not month:
        raise ValueError("Não consegui identificar o mês pelo nome da guia selecionada.")

    base_rows = read_esocial_base_rows(base_file, base_sheet)

    export_frames = []
    source_notes = []
    total_files = max(1, len(export_files))
    for idx, file_path in enumerate(export_files, start=1):
        if progress:
            progress(20 + int((idx - 1) * 30 / total_files), f"Lendo planilha eSocial {idx}/{total_files}: {Path(file_path).name}")
        frame, sheet = read_esocial_export_file(file_path)
        export_frames.append(frame)
        source_notes.append(f"{Path(file_path).name} -> guia {sheet}")

    if not export_frames:
        raise ValueError("Envie pelo menos uma planilha de envios do eSocial.")

    exports = pd.concat(export_frames, ignore_index=True)
    # Remove linhas repetidas quando o mesmo arquivo/registro é enviado mais de uma vez.
    exports = exports.drop_duplicates(
        subset=["CNPJ", "NOME_KEY", "DATA_REF", "RECIBO", "EVENTO"],
        keep="first",
    ).reset_index(drop=True)

    # Processa somente CNPJs realmente presentes nas planilhas enviadas pelo usuário.
    company_order = list(dict.fromkeys(exports["CNPJ"].dropna().astype(str).tolist()))
    generated = []
    summary = []

    os.makedirs(output_folder, exist_ok=True)
    total_companies = max(1, len(company_order))
    for pos, cnpj in enumerate(company_order, start=1):
        if progress:
            progress(52 + int((pos - 1) * 35 / total_companies), f"Cruzando empresa {pos}/{total_companies}...")
        export_company = exports[exports["CNPJ"] == cnpj].copy()
        base_company = base_rows[base_rows["CNPJ"] == cnpj].copy()
        export_name = _safe_text(export_company.iloc[0].get("EMPRESA", "")) if not export_company.empty else ""

        if base_company.empty:
            summary.append({
                "empresa": export_name or "EMPRESA NÃO IDENTIFICADA",
                "cnpj": cnpj,
                "status": "NÃO GERADO",
                "total_base": 0,
                "total_export": len(export_company),
                "total_encontrado": 0,
                "motivo": "CNPJ não possui linhas com OK E-SOCIAL na guia selecionada.",
                "pdf": "",
                "faltantes": [],
            })
            continue

        company_name = _safe_text(base_company.iloc[0]["EMPRESA_NOME"])
        selected, missing = select_esocial_rows_for_company(base_company, export_company)
        if selected.empty:
            summary.append({
                "empresa": company_name,
                "cnpj": cnpj,
                "status": "NÃO GERADO",
                "total_base": len(base_company),
                "total_export": len(export_company),
                "total_encontrado": 0,
                "motivo": "Nenhum funcionário da planilha enviada coincide com a planilha base.",
                "pdf": "",
                "faltantes": missing,
            })
            continue

        pdf_name = build_esocial_pdf_filename(company_name, cnpj, month)
        pdf_path = os.path.join(output_folder, pdf_name)
        build_esocial_receipt_pdf(selected, pdf_path, company_name, cnpj)
        generated.append(pdf_path)
        summary.append({
            "empresa": company_name,
            "cnpj": cnpj,
            "status": "GERADO",
            "total_base": len(base_company),
            "total_export": len(export_company),
            "total_encontrado": len(selected),
            "motivo": "OK",
            "pdf": pdf_name,
            "faltantes": missing,
        })

    # Mesmo quando nenhuma empresa possui funcionários em comum, o processamento
    # deve terminar normalmente. Nesse cenário o ZIP conterá o resumo detalhado,
    # permitindo ao usuário conferir quais nomes da base não estavam na exportação
    # do eSocial, em vez de receber um erro genérico.
    summary_path = os.path.join(output_folder, "RESUMO PROCESSAMENTO.txt")
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write("RECIBO eSOCIAL - RESUMO DO PROCESSAMENTO\n")
        f.write("=" * 78 + "\n")
        f.write(f"Guia base: {base_sheet}\n")
        f.write(f"Mês dos PDFs: {month}\n")
        if year:
            f.write(f"Ano identificado: {year}\n")
        f.write(f"Linhas com OK E-SOCIAL válidas na base: {len(base_rows)}\n")
        f.write(f"Planilhas de eSocial recebidas: {len(export_files)}\n")
        for note in source_notes:
            f.write(f"  - {note}\n")
        f.write("\nEMPRESAS\n" + "-" * 78 + "\n")
        for item in summary:
            f.write(f"{item['empresa']} | {format_cnpj(item['cnpj'])} | {item['status']} | ")
            f.write(
                f"base={item['total_base']} | exportação={item.get('total_export', 0)} | "
                f"encontrados={item['total_encontrado']} | {item['motivo']}\n"
            )
            if item.get("faltantes"):
                f.write("  Sem correspondência na exportação: " + "; ".join(item["faltantes"]) + "\n")

    return {
        "month": month,
        "year": year,
        "generated": generated,
        "summary": summary,
        "summary_path": summary_path,
        "total_generated": len(generated),
        "total_companies": len(company_order),
        "total_without_match": sum(1 for item in summary if item.get("status") != "GERADO"),
    }


def run_company_process(system_file: str, base_file: str, pdf_folder: str, log_folder: str, pdf_month: str, base_sheet: str | None = None):
    """Wrapper legado. O novo fluxo processa todas as planilhas juntas em process_esocial_receipts."""
    result = process_esocial_receipts(base_file, base_sheet or "", [system_file], pdf_folder)
    item = result["summary"][0] if result["summary"] else {}
    return item

# =========================
# FÍSICO E MENTAL
# =========================
def fisico_get_conn():
    conn = sqlite3.connect(FISICO_DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_fisico_db():
    with fisico_get_conn() as conn:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS fisico_empresas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS fisico_cargos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        cargo_cols = [row['name'] for row in conn.execute("PRAGMA table_info(fisico_cargos)").fetchall()]
        if 'empresa_id' in cargo_cols:
            conn.execute('ALTER TABLE fisico_cargos RENAME TO fisico_cargos_old')
            conn.execute('''
                CREATE TABLE fisico_cargos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL UNIQUE,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            conn.execute('''
                INSERT OR IGNORE INTO fisico_cargos (nome, created_at)
                SELECT DISTINCT nome, COALESCE(created_at, CURRENT_TIMESTAMP)
                FROM fisico_cargos_old
                WHERE nome IS NOT NULL AND TRIM(nome) <> ''
            ''')
            conn.execute('DROP TABLE fisico_cargos_old')
        conn.commit()

def fisico_clean_text(value: str) -> str:
    value = (value or '').strip()
    value = re.sub(r'\s+', ' ', value)
    return value.upper()

def fisico_slugify(value: str) -> str:
    normalized = unicodedata.normalize('NFKD', value or '')
    ascii_text = normalized.encode('ascii', 'ignore').decode('ascii')
    ascii_text = re.sub(r'[^A-Za-z0-9]+', '_', ascii_text).strip('_')
    return ascii_text or 'documento'

def fisico_make_rich(value: str) -> RichText:
    rt = RichText()
    rt.add(fisico_clean_text(value), bold=True)
    return rt

def docx_escape_text(value: str) -> str:
    value = '' if value is None else str(value)
    lines = value.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    escaped = [line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;') for line in lines]
    return '</w:t><w:br/><w:t>'.join(escaped)


def docx_placeholder_pattern(placeholder: str) -> str:
    # O Word pode quebrar um placeholder entre vários runs/tags XML.
    return r'(?:<[^>]+>)*'.join(re.escape(ch) for ch in placeholder)


def replace_docx_placeholders_preserve_layout(template_path: str, output_path: str, replacements: dict[str, str]) -> None:
    """Substitui placeholders em documento, cabeçalhos e rodapés sem remontar o layout."""
    with zipfile.ZipFile(template_path, 'r') as zin, zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml' or item.filename.startswith('word/header') or item.filename.startswith('word/footer'):
                try:
                    xml = data.decode('utf-8')
                except UnicodeDecodeError:
                    zout.writestr(item, data)
                    continue
                for placeholder, value in replacements.items():
                    xml = re.sub(
                        docx_placeholder_pattern(placeholder),
                        lambda _m, v=value: docx_escape_text(v),
                        xml,
                    )
                data = xml.encode('utf-8')
            zout.writestr(item, data)


def fisico_build_orgao_texto(empresa: str, edital: str, pss: str) -> str:
    parts = [fisico_clean_text(empresa), fisico_clean_text(edital), fisico_clean_text(pss)]
    parts = [p for p in parts if p]
    return ' - '.join(parts)

def fisico_format_date_extenso(raw_date: str) -> str:
    if raw_date:
        dt = datetime.strptime(raw_date, '%Y-%m-%d').date()
    else:
        dt = datetime.today().date()
    return f'{dt.day:02d} DE {nome_mes(dt.month)} DE {dt.year}'

def fisico_convert_to_pdf(docx_path: str, target_dir: str) -> str:
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        raise RuntimeError('LibreOffice/soffice não encontrado para converter PDF.')
    cmd = [soffice, '--headless', '--convert-to', 'pdf', '--outdir', target_dir, docx_path]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or 'Falha ao converter para PDF.')
    pdf_path = os.path.join(target_dir, f"{Path(docx_path).stem}.pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError('PDF não foi gerado.')
    return pdf_path

def fisico_list_empresas(search: str = ''):
    query = 'SELECT id, nome FROM fisico_empresas'
    params = []
    if search:
        query += ' WHERE nome LIKE ?'
        params.append(f'%{search}%')
    query += ' ORDER BY nome'
    with fisico_get_conn() as conn:
        return conn.execute(query, params).fetchall()

def fisico_list_cargos(search: str = ''):
    query = 'SELECT id, nome FROM fisico_cargos'
    params = []
    if search:
        query += ' WHERE nome LIKE ?'
        params.append(f'%{search}%')
    query += ' ORDER BY nome'
    with fisico_get_conn() as conn:
        return conn.execute(query, params).fetchall()

def fisico_render_home(form_data=None):
    form_data = form_data or {}
    return render_template('fisico_mental.html',
                           title='Físico e Mental',
                           today=form_data.get('data_exame') or datetime.today().strftime('%Y-%m-%d'),
                           empresas=fisico_list_empresas(),
                           cargos=fisico_list_cargos(),
                           locais=CLINIC_LOCATIONS,
                           form_data=form_data)

def fisico_render_cadastros(search=''):
    return render_template('fisico_mental_cadastros.html',
                           title='Cadastros Físico e Mental',
                           search=search,
                           empresas=fisico_list_empresas(search),
                           cargos=fisico_list_cargos(search))

init_fisico_db()

# =========================
# ROTAS
# =========================
@app.route('/fisico-mental', methods=['GET'])
def fisico_mental():
    return fisico_render_home()

@app.route('/fisico-mental/cadastros', methods=['GET'])
def fisico_mental_cadastros():
    search = (request.args.get('q') or '').strip()
    return fisico_render_cadastros(search)

@app.route('/fisico-mental/cadastros/empresa/adicionar', methods=['POST'])
def fisico_adicionar_empresa():
    nome = fisico_clean_text(request.form.get('nome', ''))
    if not nome:
        flash('Digite o nome da empresa.', 'error')
        return redirect(url_for('fisico_mental_cadastros'))
    try:
        with fisico_get_conn() as conn:
            conn.execute('INSERT INTO fisico_empresas (nome) VALUES (?)', (nome,))
            conn.commit()
        flash('Empresa cadastrada com sucesso.', 'success')
    except sqlite3.IntegrityError:
        flash('Essa empresa já está cadastrada.', 'error')
    return redirect(url_for('fisico_mental_cadastros'))

@app.route('/fisico-mental/cadastros/empresa/editar', methods=['POST'])
def fisico_editar_empresa():
    item_id = request.form.get('id', '')
    nome = fisico_clean_text(request.form.get('nome', ''))
    if not item_id.isdigit() or not nome:
        flash('Não foi possível editar a empresa.', 'error')
        return redirect(url_for('fisico_mental_cadastros'))
    try:
        with fisico_get_conn() as conn:
            conn.execute('UPDATE fisico_empresas SET nome = ? WHERE id = ?', (nome, int(item_id)))
            conn.commit()
        flash('Empresa atualizada com sucesso.', 'success')
    except sqlite3.IntegrityError:
        flash('Já existe outra empresa com esse nome.', 'error')
    return redirect(url_for('fisico_mental_cadastros'))

@app.route('/fisico-mental/cadastros/empresa/excluir', methods=['POST'])
def fisico_excluir_empresa():
    item_id = request.form.get('id', '')
    if not item_id.isdigit():
        flash('Não foi possível excluir a empresa.', 'error')
        return redirect(url_for('fisico_mental_cadastros'))
    with fisico_get_conn() as conn:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.execute('DELETE FROM fisico_empresas WHERE id = ?', (int(item_id),))
        conn.commit()
    flash('Empresa excluída.', 'success')
    return redirect(url_for('fisico_mental_cadastros'))

@app.route('/fisico-mental/cadastros/cargo/adicionar', methods=['POST'])
def fisico_adicionar_cargo():
    nome = fisico_clean_text(request.form.get('nome', ''))
    if not nome:
        flash('Digite o nome do cargo.', 'error')
        return redirect(url_for('fisico_mental_cadastros'))
    try:
        with fisico_get_conn() as conn:
            conn.execute('INSERT INTO fisico_cargos (nome) VALUES (?)', (nome,))
            conn.commit()
        flash('Cargo cadastrado com sucesso.', 'success')
    except sqlite3.IntegrityError:
        flash('Esse cargo já está cadastrado.', 'error')
    return redirect(url_for('fisico_mental_cadastros'))

@app.route('/fisico-mental/cadastros/cargo/editar', methods=['POST'])
def fisico_editar_cargo():
    item_id = request.form.get('id', '')
    nome = fisico_clean_text(request.form.get('nome', ''))
    if not item_id.isdigit() or not nome:
        flash('Não foi possível editar o cargo.', 'error')
        return redirect(url_for('fisico_mental_cadastros'))
    try:
        with fisico_get_conn() as conn:
            conn.execute('UPDATE fisico_cargos SET nome = ? WHERE id = ?', (nome, int(item_id)))
            conn.commit()
        flash('Cargo atualizado com sucesso.', 'success')
    except sqlite3.IntegrityError:
        flash('Já existe esse cargo cadastrado.', 'error')
    return redirect(url_for('fisico_mental_cadastros'))

@app.route('/fisico-mental/cadastros/cargo/excluir', methods=['POST'])
def fisico_excluir_cargo():
    item_id = request.form.get('id', '')
    if not item_id.isdigit():
        flash('Não foi possível excluir o cargo.', 'error')
        return redirect(url_for('fisico_mental_cadastros'))
    with fisico_get_conn() as conn:
        conn.execute('DELETE FROM fisico_cargos WHERE id = ?', (int(item_id),))
        conn.commit()
    flash('Cargo excluído.', 'success')
    return redirect(url_for('fisico_mental_cadastros'))

@app.route('/fisico-mental/gerar', methods=['POST'])
def fisico_gerar():
    form_data = request.form.to_dict(flat=True)
    nome = fisico_clean_text(request.form.get('nome', ''))
    rg = fisico_clean_text(request.form.get('rg', ''))
    cpf = fisico_clean_text(request.form.get('cpf', ''))
    empresa = fisico_clean_text(request.form.get('empresa_nome', ''))
    edital = fisico_clean_text(request.form.get('edital', ''))
    pss = fisico_clean_text(request.form.get('pss', ''))
    funcao = fisico_clean_text(request.form.get('funcao_nome', ''))
    data_exame = request.form.get('data_exame', '')
    local_key = (request.form.get('local_exame') or '').strip().lower()
    formato = (request.form.get('formato', 'docx') or 'docx').lower()
    local = CLINIC_LOCATIONS.get(local_key)

    if not nome or not rg or not cpf or not empresa or not funcao or not local:
        flash('Preencha nome, RG, CPF, empresa, cargo e unidade de atendimento.', 'error')
        return fisico_render_home(form_data)

    if len(somente_numeros(cpf)) != 11:
        flash('CPF inválido. Informe os 11 números do CPF.', 'error')
        return fisico_render_home(form_data)

    replacements = {
        '{{ nome }}': nome,
        '{{nome}}': nome,
        '{{ rg }}': rg,
        '{{rg}}': rg,
        '{{ cpf }}': cpf,
        '{{cpf}}': cpf,
        '{{ empresa }}': empresa,
        '{{empresa}}': empresa,
        '{{ orgao_texto }}': fisico_build_orgao_texto(empresa, edital, pss),
        '{{orgao_texto}}': fisico_build_orgao_texto(empresa, edital, pss),
        '{{ funcao }}': funcao,
        '{{funcao}}': funcao,
        '{{ data_extenso }}': fisico_format_date_extenso(data_exame),
        '{{data_extenso}}': fisico_format_date_extenso(data_exame),
        '{{CIDADE, UF}}': local['cidade_uf'],
        '{{ENDEREÇO}}': local['endereco'],
        '{{telefone}}': local['telefone'],
        '{{informações do dr}}': local['medico_fisico_mental'],
    }

    filename_base = fisico_slugify(f'fisico_mental_{nome}')
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, f'{filename_base}.docx')
        try:
            replace_docx_placeholders_preserve_layout(FISICO_TEMPLATE_PATH, docx_path, replacements)
        except Exception:
            logger.exception('Erro ao gerar atestado físico e mental')
            flash('Não foi possível gerar o documento. Confira os dados e tente novamente.', 'error')
            return fisico_render_home(form_data)
        if formato == 'pdf':
            try:
                pdf_path = fisico_convert_to_pdf(docx_path, tmpdir)
                payload = Path(pdf_path).read_bytes()
                return send_file(io.BytesIO(payload), as_attachment=True, download_name=f'{filename_base}.pdf', mimetype='application/pdf')
            except Exception as exc:
                flash(f'Não foi possível gerar PDF agora: {exc}. O arquivo foi enviado em Word.', 'error')
        payload = Path(docx_path).read_bytes()
        return send_file(io.BytesIO(payload), as_attachment=True, download_name=f'{filename_base}.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


# =========================
# ATESTADO MÉDICO
# =========================
ATESTADO_CIDADES = {key: data['label'] for key, data in CLINIC_LOCATIONS.items()}


def atestado_render_home(form_data=None):
    form_data = form_data or {}
    return render_template(
        'atestado_medico.html',
        title='Atestado Médico',
        today=form_data.get('data') or datetime.today().strftime('%Y-%m-%d'),
        cidades=ATESTADO_CIDADES,
        form_data=form_data,
    )


def atestado_format_date(raw_date: str) -> str:
    if raw_date:
        try:
            dt = datetime.strptime(raw_date, '%Y-%m-%d').date()
        except Exception:
            dt = datetime.today().date()
    else:
        dt = datetime.today().date()
    return dt.strftime('%d/%m/%Y')


def atestado_docx_escape(value: str) -> str:
    value = '' if value is None else str(value)
    lines = value.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    escaped = [line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;') for line in lines]
    return '</w:t><w:br/><w:t>'.join(escaped)


def atestado_placeholder_pattern(placeholder: str) -> str:
    return r'(?:<[^>]+>)*'.join(re.escape(ch) for ch in placeholder)


def atestado_replace_placeholder(xml: str, placeholder: str, value: str) -> str:
    pattern = atestado_placeholder_pattern(placeholder)
    return re.sub(pattern, lambda _match: atestado_docx_escape(value), xml)


def atestado_replace_docx_placeholders(template_path: str, output_path: str, replacements: dict[str, str]) -> None:
    """Substitui os campos do modelo preservando layout, imagens, tabelas e rodapé."""
    xml_targets = ('word/document.xml', 'word/header', 'word/footer')
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml' or item.filename.startswith(xml_targets[1]) or item.filename.startswith(xml_targets[2]):
                    try:
                        xml = data.decode('utf-8')
                    except UnicodeDecodeError:
                        zout.writestr(item, data)
                        continue
                    for placeholder, value in replacements.items():
                        xml = atestado_replace_placeholder(xml, placeholder, value)
                    data = xml.encode('utf-8')
                zout.writestr(item, data)


@app.route('/atestado-medico', methods=['GET'])
def atestado_medico():
    return atestado_render_home()


@app.route('/atestado-medico/gerar', methods=['POST'])
def atestado_medico_gerar():
    form_data = request.form.to_dict(flat=True)
    nome = fisico_clean_text(request.form.get('nome', ''))
    data_atestado = request.form.get('data', '')
    dias = (request.form.get('dias') or '').strip()
    cid = fisico_clean_text(request.form.get('cid', ''))
    cidade_key = request.form.get('cidade_uf', '')
    formato = (request.form.get('formato', 'docx') or 'docx').lower()

    if not nome or not data_atestado or not dias or not cid or not cidade_key:
        flash('Preencha nome, data, quantidade de dias, CID e cidade/UF.', 'error')
        return atestado_render_home(form_data)

    if cidade_key not in CLINIC_LOCATIONS:
        flash('Selecione uma unidade válida: BELÉM, PA ou MACAPÁ, AP.', 'error')
        return atestado_render_home(form_data)

    if not dias.isdigit() or int(dias) <= 0:
        flash('Informe a quantidade de dias usando apenas números.', 'error')
        return atestado_render_home(form_data)

    data_formatada = atestado_format_date(data_atestado)
    local = CLINIC_LOCATIONS[cidade_key]
    replacements = {
        '{{Nome}}': nome,
        '{{data}}': data_formatada,
        '{{dias}}': str(int(dias)),
        '{{cid}}': cid,
        '{{cidade-uf}}': local['cidade_hifen'],
        '{{CIDADE, UF}}': local['cidade_uf'],
        '{{ENDEREÇO}}': local['endereco'],
        '{{telefone}}': local['telefone'],
    }

    filename_base = sanitize_filename(f'ATESTADO MEDICO - {nome}')
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, f'{filename_base}.docx')
        try:
            atestado_replace_docx_placeholders(ATESTADO_MEDICO_TEMPLATE_PATH, docx_path, replacements)
        except Exception:
            logger.exception('Erro ao gerar atestado médico')
            flash('Não foi possível gerar o atestado. Confira os dados e tente novamente.', 'error')
            return atestado_render_home(form_data)

        if formato == 'pdf':
            try:
                pdf_path = fisico_convert_to_pdf(docx_path, tmpdir)
                payload = Path(pdf_path).read_bytes()
                return send_file(io.BytesIO(payload), as_attachment=True, download_name=f'{filename_base}.pdf', mimetype='application/pdf')
            except Exception as exc:
                flash(f'Não foi possível gerar PDF agora: {exc}. O arquivo foi enviado em Word.', 'error')

        payload = Path(docx_path).read_bytes()
        return send_file(io.BytesIO(payload), as_attachment=True, download_name=f'{filename_base}.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')



# =========================
# ANAMNESE OCUPACIONAL - SRQ-20
# =========================
def anamnese_render_home(form_data=None):
    form_data = form_data or {}
    return render_template(
        'anamnese_ocupacional.html',
        title='Anamnese Ocupacional',
        form_data=form_data,
    )


def anamnese_fill_docx(template_path: str, output_path: str, empresa: str, cnpj: str, nome: str, funcao: str) -> None:
    """Preenche somente os campos explicitamente marcados com {{ }} no modelo.

    As perguntas, colunas SIM/NÃO e a pontuação permanecem em branco para
    preenchimento manual no documento impresso.
    """
    replacements = {
        '{{EMPRESA}}': empresa,
        '{{CNPJ}}': cnpj,
        '{{NOME}}': nome,
        '{{FUNCAO}}': funcao,
    }
    replace_docx_placeholders_preserve_layout(template_path, output_path, replacements)


@app.route('/anamnese-ocupacional', methods=['GET'])
def anamnese_ocupacional():
    return anamnese_render_home()


@app.route('/anamnese-ocupacional/gerar', methods=['POST'])
def anamnese_ocupacional_gerar():
    form_data = request.form.to_dict(flat=True)
    empresa = fisico_clean_text(request.form.get('empresa', ''))
    cnpj = (request.form.get('cnpj') or '').strip()
    nome = fisico_clean_text(request.form.get('nome', ''))
    funcao = fisico_clean_text(request.form.get('funcao', ''))
    formato = (request.form.get('formato', 'docx') or 'docx').lower()

    if not empresa or not cnpj or not nome or not funcao:
        flash('Preencha empresa, CNPJ, nome e função.', 'error')
        return anamnese_render_home(form_data)

    cnpj_digits = somente_numeros(cnpj)
    if len(cnpj_digits) != 14:
        flash('CNPJ inválido. Informe os 14 números do CNPJ.', 'error')
        return anamnese_render_home(form_data)

    filename_base = sanitize_filename(f'ANAMNESE OCUPACIONAL - {nome}')
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, f'{filename_base}.docx')
        try:
            anamnese_fill_docx(
                ANAMNESE_OCUPACIONAL_TEMPLATE_PATH,
                docx_path,
                empresa,
                cnpj,
                nome,
                funcao,
            )
        except Exception:
            logger.exception('Erro ao gerar anamnese ocupacional')
            flash('Não foi possível gerar a anamnese. Confira os dados e tente novamente.', 'error')
            return anamnese_render_home(form_data)

        if formato == 'pdf':
            try:
                pdf_path = fisico_convert_to_pdf(docx_path, tmpdir)
                payload = Path(pdf_path).read_bytes()
                return send_file(io.BytesIO(payload), as_attachment=True, download_name=f'{filename_base}.pdf', mimetype='application/pdf')
            except Exception as exc:
                flash(f'Não foi possível gerar PDF agora: {exc}. O arquivo foi enviado em Word.', 'error')

        payload = Path(docx_path).read_bytes()
        return send_file(
            io.BytesIO(payload),
            as_attachment=True,
            download_name=f'{filename_base}.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )


# =========================
# ASO MANUAL
# =========================
def aso_today_br() -> str:
    """Data local padrão das unidades EDGE (Belém/Macapá, UTC-3) em DD/MM/AAAA."""
    return (datetime.utcnow() - timedelta(hours=3)).strftime('%d/%m/%Y')


def aso_today_iso() -> str:
    """Data local padrão para campos HTML date (AAAA-MM-DD)."""
    return (datetime.utcnow() - timedelta(hours=3)).strftime('%Y-%m-%d')


def aso_manual_render_home(form_data=None):
    form_data = dict(form_data or {})
    return render_template(
        'aso_manual.html',
        title='ASO manual',
        locais=CLINIC_LOCATIONS,
        form_data=form_data,
        today_br=aso_today_br(),
        today_iso=aso_today_iso(),
    )


def parse_date_br(raw_date: str):
    """Aceita DD/MM/AAAA, DD-MM-AAAA, DD.MM.AAAA, AAAA-MM-DD e datas coladas só com números."""
    value = (raw_date or '').strip()
    if not value:
        return None

    normalized = value.replace('.', '/').replace('-', '/')
    for fmt in ('%d/%m/%Y', '%Y/%m/%d'):
        try:
            return datetime.strptime(normalized, fmt)
        except ValueError:
            pass

    digits = re.sub(r'\D', '', value)
    if len(digits) == 8:
        # AAAAMMDD quando os quatro primeiros dígitos formam um ano plausível;
        # caso contrário, interpreta como DDMMAAAA.
        year_first = int(digits[:4])
        formats = ('%Y%m%d', '%d%m%Y') if 1900 <= year_first <= 2100 else ('%d%m%Y', '%Y%m%d')
        for fmt in formats:
            try:
                return datetime.strptime(digits, fmt)
            except ValueError:
                pass

    raise ValueError('Data inválida.')


def format_date_br(raw_date: str) -> str:
    """Converte datas dos formulários para DD/MM/AAAA; vazio permanece vazio."""
    parsed = parse_date_br(raw_date)
    return parsed.strftime('%d/%m/%Y') if parsed else ''


def calculate_age_from_birth(birth_date, reference_date) -> int:
    if birth_date is None:
        raise ValueError('Data de nascimento inválida.')
    reference_date = reference_date or datetime.utcnow()
    if birth_date.date() > reference_date.date():
        raise ValueError('A data de nascimento não pode ser posterior à data de referência.')
    return reference_date.year - birth_date.year - (
        (reference_date.month, reference_date.day) < (birth_date.month, birth_date.day)
    )


def aso_set_cell_text(cell, text: str) -> None:
    """Atualiza o texto de uma célula preservando o máximo do estilo/layout original."""
    paragraph = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    if paragraph.runs:
        first = paragraph.runs[0]
        first.text = text
        for run in paragraph.runs[1:]:
            run.text = ''
    else:
        paragraph.add_run(text)


def aso_fill_docx(template_path: str, output_path: str, replacements: dict[str, str], complementares: list[tuple[str, str]]) -> None:
    """Preenche o ASO preservando exatamente o espaçamento e tamanho do modelo original."""
    temp_path = output_path + '.base.docx'
    replace_docx_placeholders_preserve_layout(template_path, temp_path, replacements)
    try:
        doc = Document(temp_path)
        if not doc.tables or len(doc.tables[0].rows) < 4:
            raise ValueError('O modelo de ASO não contém a tabela de exames esperada.')
        table = doc.tables[0]
        # Visualmente: 1/4 na primeira linha, 2/5 na segunda, 3/6 na terceira.
        mapping = [
            (1, 0, 1, 1),
            (2, 0, 2, 1),
            (3, 0, 3, 1),
            (1, 2, 1, 3),
            (2, 2, 2, 3),
            (3, 2, 3, 3),
        ]
        for idx, (row_exam, col_exam, row_date, col_date) in enumerate(mapping, start=1):
            exame, data = complementares[idx - 1]
            if idx == 1:
                exam_text = '1 – EXAME CLÍNICO'
            else:
                exam_text = f'{idx} – {exame}' if exame else f'{idx} –'
            aso_set_cell_text(table.rows[row_exam].cells[col_exam], exam_text)
            aso_set_cell_text(table.rows[row_date].cells[col_date], data)

        doc.save(output_path)
    finally:
        try:
            os.remove(temp_path)
        except OSError:
            pass


@app.route('/aso-manual', methods=['GET'])
def aso_manual():
    return aso_manual_render_home()


@app.route('/aso-manual/gerar', methods=['POST'])
def aso_manual_gerar():
    form_data = request.form.to_dict(flat=True)
    unidade = (request.form.get('unidade') or '').strip().lower()
    local = CLINIC_LOCATIONS.get(unidade)
    if not local:
        flash('Selecione a unidade de Belém ou Macapá.', 'error')
        return aso_manual_render_home(form_data)

    empresa = fisico_clean_text(request.form.get('empresa', ''))
    cnpj = (request.form.get('cnpj') or '').strip()
    funcionario = fisico_clean_text(request.form.get('funcionario', ''))
    rg = fisico_clean_text(request.form.get('rg', ''))
    cpf = (request.form.get('cpf') or '').strip()
    data_nascimento_raw = (request.form.get('data_nascimento') or '').strip()
    cargo = fisico_clean_text(request.form.get('cargo', ''))
    setor = fisico_clean_text(request.form.get('setor', ''))
    tipo_exame = fisico_clean_text(request.form.get('tipo_exame', ''))
    data_aso_raw = (request.form.get('data_aso') or '').strip()
    data_clinico_raw = (request.form.get('datacomp_1') or '').strip()
    formato = (request.form.get('formato', 'docx') or 'docx').lower()

    required_fields = [
        ('Empresa', empresa),
        ('CNPJ da empresa', cnpj),
        ('Funcionário', funcionario),
        ('RG', rg),
        ('CPF', cpf),
        ('Data de nascimento', data_nascimento_raw),
        ('Cargo', cargo),
        ('Setor', setor),
        ('Tipo de exame', tipo_exame),
        ('Data do ASO', data_aso_raw),
        ('Data do Exame Clínico', data_clinico_raw),
    ]
    missing = [label for label, value in required_fields if not value]
    if missing:
        flash('Preencha antes de gerar: ' + ', '.join(missing) + '.', 'error')
        return aso_manual_render_home(form_data)

    if len(somente_numeros(cnpj)) != 14:
        flash('CNPJ inválido. Informe os 14 números do CNPJ.', 'error')
        return aso_manual_render_home(form_data)
    if len(somente_numeros(cpf)) != 11:
        flash('CPF inválido. Informe os 11 números do CPF.', 'error')
        return aso_manual_render_home(form_data)

    try:
        nascimento_dt = parse_date_br(data_nascimento_raw)
        data_nascimento = nascimento_dt.strftime('%d/%m/%Y')
        data_aso_dt = parse_date_br(data_aso_raw)
        data_aso = data_aso_dt.strftime('%d/%m/%Y')
        idade = str(calculate_age_from_birth(nascimento_dt, data_aso_dt))

        # Exame clínico é fixo; sua data é obrigatória e vem com hoje por padrão na tela.
        complementares = [('EXAME CLÍNICO', format_date_br(data_clinico_raw))]
        for numero in range(2, 7):
            exame = fisico_clean_text(request.form.get(f'complementar_{numero}', ''))
            data_raw = (request.form.get(f'datacomp_{numero}') or '').strip()
            if exame and not data_raw:
                flash(f'Informe a data do Exame {numero} ({exame}) antes de gerar o ASO.', 'error')
                return aso_manual_render_home(form_data)
            if data_raw and not exame:
                flash(f'Informe o nome do Exame {numero} ou apague a data preenchida.', 'error')
                return aso_manual_render_home(form_data)
            data_comp = format_date_br(data_raw) if data_raw else ''
            complementares.append((exame, data_comp))
    except ValueError as exc:
        flash(str(exc) if str(exc) else 'Confira as datas informadas.', 'error')
        return aso_manual_render_home(form_data)

    replacements = {
        '{{cnpjedge}}': local['cnpj_edge_aso'],
        '{{EMPRESA}}': empresa,
        '{{CNPJ}}': cnpj,
        '{{FUNCIONARIO}}': funcionario,
        '{{FUNCIONÁRIO}}': funcionario,
        '{{RG}}': rg,
        '{{CPF}}': cpf,
        '{{DATANASC}}': data_nascimento,
        '{{DATANASCIMENTO}}': data_nascimento,
        '{{IDADE}}': idade,
        '{{}}': idade,
        '{{CARGO}}': cargo,
        '{{SETOR}}': setor,
        '{{TIPODEEXAME}}': tipo_exame,
        '{{DATA}}': data_aso,
        '{{INFODOUTOR}}': local['info_doutor_aso'],
        '{{UF}}': local['uf'],
    }

    filename_base = sanitize_filename(f'ASO MANUAL - {funcionario}')
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, f'{filename_base}.docx')
        try:
            aso_fill_docx(ASO_MANUAL_TEMPLATE_PATH, docx_path, replacements, complementares)
        except Exception:
            logger.exception('Erro ao gerar ASO manual')
            flash('Não foi possível gerar o ASO manual. Confira os dados e tente novamente.', 'error')
            return aso_manual_render_home(form_data)

        if formato == 'pdf':
            try:
                pdf_path = fisico_convert_to_pdf(docx_path, tmpdir)
                payload = Path(pdf_path).read_bytes()
                return send_file(io.BytesIO(payload), as_attachment=True, download_name=f'{filename_base}.pdf', mimetype='application/pdf')
            except Exception as exc:
                flash(f'Não foi possível gerar PDF agora: {exc}. O arquivo foi enviado em Word.', 'error')

        payload = Path(docx_path).read_bytes()
        return send_file(
            io.BytesIO(payload),
            as_attachment=True,
            download_name=f'{filename_base}.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )

# =========================
# LAUDO PCD
# =========================
PCD_TIPOS = {
    'fisica': 'FÍSICA',
    'auditiva': 'AUDITIVA',
    'visual': 'VISUAL',
    'intelectual': 'INTELECTUAL/MENTAL',
    'multipla': 'MÚLTIPLA',
}


def pcd_render_home(form_data=None):
    form_data = form_data or {}
    return render_template(
        'laudo_pcd.html',
        title='Laudo PCD',
        today=form_data.get('data_laudo') or datetime.today().strftime('%Y-%m-%d'),
        empresas=fisico_list_empresas(),
        cargos=fisico_list_cargos(),
        tipos=PCD_TIPOS,
        locais=CLINIC_LOCATIONS,
        form_data=form_data,
    )


def pcd_docx_escape(value: str) -> str:
    value = '' if value is None else str(value)
    lines = value.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    escaped = [line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;') for line in lines]
    return '</w:t><w:br/><w:t>'.join(escaped)


def pcd_placeholder_pattern(placeholder: str) -> str:
    """Regex que encontra placeholder mesmo quando o Word divide em vários runs."""
    return r'(?:<[^>]+>)*'.join(re.escape(ch) for ch in placeholder)


def pcd_replace_placeholder(xml: str, placeholder: str, value: str, count: int = 0) -> str:
    pattern = pcd_placeholder_pattern(placeholder)
    return re.sub(pattern, lambda _match: pcd_docx_escape(value), xml, count=count)


def pcd_format_date_parts(raw_date: str) -> tuple[str, str, str]:
    if raw_date:
        try:
            dt = datetime.strptime(raw_date, '%Y-%m-%d').date()
        except Exception:
            dt = datetime.today().date()
    else:
        dt = datetime.today().date()
    return f'{dt.day:02d}', nome_mes(dt.month), str(dt.year)


def pcd_replace_docx_placeholders(template_path: str, output_path: str, replacements: dict[str, str], sequence_replacements: dict[str, list[str]]):
    """Substitui placeholders do modelo PCD preservando o pacote DOCX original.

    O modelo recebido possui campos como {{EMPRESA - CNPJ}}, que não são nomes
    válidos para Jinja/docxtpl. Por isso a substituição é feita diretamente no
    XML do Word, mantendo layout, tabelas e imagens do arquivo enviado.
    """
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    xml = data.decode('utf-8')
                    auditiva_mark = replacements.get('__AUDITIVA_MARK__', '')
                    xml = re.sub(r'<w:t>XX</w:t>', f'<w:t>{pcd_docx_escape(auditiva_mark)}</w:t>', xml)
                    for placeholder, values in sequence_replacements.items():
                        for value in values:
                            xml = pcd_replace_placeholder(xml, placeholder, value, count=1)
                        xml = pcd_replace_placeholder(xml, placeholder, '')
                    for placeholder, value in replacements.items():
                        if placeholder.startswith('__'):
                            continue
                        xml = pcd_replace_placeholder(xml, placeholder, value)
                    data = xml.encode('utf-8')
                zout.writestr(item, data)


def pcd_quick_add_empresa(nome: str) -> tuple[bool, str, dict]:
    nome = fisico_clean_text(nome)
    if not nome:
        return False, 'Informe o nome da empresa.', {}
    with fisico_get_conn() as conn:
        existente = conn.execute('SELECT id, nome FROM fisico_empresas WHERE nome = ?', (nome,)).fetchone()
        if existente:
            return True, 'Empresa já cadastrada.', {'id': existente['id'], 'nome': existente['nome']}
        conn.execute('INSERT INTO fisico_empresas (nome) VALUES (?)', (nome,))
        conn.commit()
        novo = conn.execute('SELECT id, nome FROM fisico_empresas WHERE nome = ?', (nome,)).fetchone()
        return True, 'Empresa cadastrada.', {'id': novo['id'], 'nome': novo['nome']}


def pcd_quick_add_cargo(nome: str) -> tuple[bool, str, dict]:
    nome = fisico_clean_text(nome)
    if not nome:
        return False, 'Informe o nome do cargo.', {}
    with fisico_get_conn() as conn:
        existente = conn.execute('SELECT id, nome FROM fisico_cargos WHERE nome = ?', (nome,)).fetchone()
        if existente:
            return True, 'Cargo já cadastrado.', {'id': existente['id'], 'nome': existente['nome']}
        conn.execute('INSERT INTO fisico_cargos (nome) VALUES (?)', (nome,))
        conn.commit()
        novo = conn.execute('SELECT id, nome FROM fisico_cargos WHERE nome = ?', (nome,)).fetchone()
        return True, 'Cargo cadastrado.', {'id': novo['id'], 'nome': novo['nome']}


@app.route('/laudo-pcd', methods=['GET'])
def laudo_pcd():
    return pcd_render_home()


@app.route('/laudo-pcd/cadastros/empresa/adicionar-rapido', methods=['POST'])
def laudo_pcd_adicionar_empresa_rapido():
    ok, msg, item = pcd_quick_add_empresa(request.form.get('nome', ''))
    return jsonify({'ok': ok, 'message': msg, 'item': item}), 200 if ok else 400


@app.route('/laudo-pcd/cadastros/cargo/adicionar-rapido', methods=['POST'])
def laudo_pcd_adicionar_cargo_rapido():
    ok, msg, item = pcd_quick_add_cargo(request.form.get('nome', ''))
    return jsonify({'ok': ok, 'message': msg, 'item': item}), 200 if ok else 400


@app.route('/laudo-pcd/gerar', methods=['POST'])
def laudo_pcd_gerar():
    form_data = request.form.to_dict(flat=True)
    tipo = (request.form.get('tipo_deficiencia') or '').strip()
    empresa = fisico_clean_text(request.form.get('empresa_nome', ''))
    empresa_cnpj = request.form.get('empresa_cnpj', '').strip()
    cargo = fisico_clean_text(request.form.get('cargo_nome', ''))
    nome = fisico_clean_text(request.form.get('nome', ''))
    cpf = request.form.get('cpf', '').strip()
    rg = fisico_clean_text(request.form.get('rg', ''))
    uf = fisico_clean_text(request.form.get('uf', ''))
    cid = fisico_clean_text(request.form.get('cid', ''))
    obs = (request.form.get('obs') or '').strip()
    data_laudo = request.form.get('data_laudo', '')
    local_key = (request.form.get('local_exame') or '').strip().lower()
    local = CLINIC_LOCATIONS.get(local_key)

    if tipo not in PCD_TIPOS:
        flash('Selecione o tipo de laudo PCD.', 'error')
        return pcd_render_home(form_data)
    if not empresa or not nome or not cpf or not rg or not cid or not local:
        flash('Preencha empresa, nome, CPF, RG, CID e unidade de atendimento.', 'error')
        return pcd_render_home(form_data)

    cpf_digits = somente_numeros(cpf)
    if len(cpf_digits) != 11:
        flash('CPF inválido. Informe os 11 números do CPF.', 'error')
        return pcd_render_home(form_data)

    dia, mes_extenso, ano = pcd_format_date_parts(data_laudo)
    empresa_doc = empresa
    if empresa_cnpj:
        empresa_doc = f'{empresa} - {formatar_documento(empresa_cnpj) or empresa_cnpj}'

    cid_values = {
        'fisica': [cid, '', '', ''],
        'auditiva': ['', cid, '', ''],
        'visual': ['', '', cid, ''],
        'intelectual': ['', '', '', cid],
        'multipla': ['', '', '', ''],
    }[tipo]
    obs_values = {
        'fisica': [obs, '', '', '', ''],
        'auditiva': ['', obs, '', '', ''],
        'visual': ['', '', obs, '', ''],
        'intelectual': ['', '', '', obs, ''],
        'multipla': ['', '', '', '', obs],
    }[tipo]

    replacements = {
        '{{EMPRESA - CNPJ}}': empresa_doc,
        '{{NOME}}': nome,
        '{{CPF}}': formatar_documento(cpf_digits) or cpf,
        '{{RG}}': rg,
        '{{UF}}': uf,
        '{{DIA}}': dia,
        '{{dia}}': dia,
        '{{CIDADE, UF}}': local['cidade_uf'],
        '{{MÊS}}': mes_extenso,
        '{{ANO}}': ano,
        '{{OlhoDAcuidade}}': request.form.get('olho_d_acuidade', ''),
        '{{OlhoEAcuidade}}': request.form.get('olho_e_acuidade', ''),
        '{{OlhoDCampo}}': request.form.get('olho_d_campo', ''),
        '{{OlhoECampo}}': request.form.get('olho_e_campo', ''),
        '{{OlhoDVisao}}': request.form.get('olho_d_visao', ''),
        '{{OlhoEVisao}}': request.form.get('olho_e_visao', ''),
        '__AUDITIVA_MARK__': 'X' if tipo == 'auditiva' else '',
    }
    sequence_replacements = {
        '{{CID}}': cid_values,
        '{{Obs}}': obs_values,
        '{{ouvidod}}': [
            request.form.get('ouvido_d_500', ''),
            request.form.get('ouvido_d_1000', ''),
            request.form.get('ouvido_d_2000', ''),
            request.form.get('ouvido_d_3000', ''),
        ],
        '{{ouvidoe}}': [
            request.form.get('ouvido_e_500', ''),
            request.form.get('ouvido_e_1000', ''),
            request.form.get('ouvido_e_2000', ''),
            request.form.get('ouvido_e_3000', ''),
        ],
    }

    filename_base = sanitize_filename(f'LAUDO PCD - {nome} - {PCD_TIPOS[tipo]}')
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = os.path.join(tmpdir, f'{filename_base}.docx')
        try:
            pcd_replace_docx_placeholders(PCD_TEMPLATE_PATH, output_path, replacements, sequence_replacements)
        except Exception:
            logger.exception('Erro ao gerar Laudo PCD')
            flash('Não foi possível gerar o Laudo PCD. Confira os dados e tente novamente.', 'error')
            return pcd_render_home(form_data)
        payload = Path(output_path).read_bytes()
        return send_file(
            io.BytesIO(payload),
            as_attachment=True,
            download_name=f'{filename_base}.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )


# =========================
# ENCAMINHAMENTO PARA ESPECIALISTA
# =========================
def encaminhamento_clean_text(value: str, upper: bool = False) -> str:
    value = (value or '').strip()
    value = re.sub(r'\s+', ' ', value)
    return value.upper() if upper else value


def encaminhamento_title_name(value: str) -> str:
    value = encaminhamento_clean_text(value)
    if not value:
        return ''
    return ' '.join(part.capitalize() for part in value.split())


def encaminhamento_empresa_name(value: str) -> str:
    return encaminhamento_clean_text(value, upper=True)


def encaminhamento_especialista_name(value: str) -> str:
    return encaminhamento_title_name(value)


ENCAMINHAMENTO_LOCAIS_EXAME = CLINIC_LOCATIONS


def encaminhamento_get_local_exame(value: str) -> dict | None:
    return ENCAMINHAMENTO_LOCAIS_EXAME.get((value or '').strip().lower())


def init_encaminhamento_db():
    """Cria os cadastros usados no encaminhamento individual.

    Usa o mesmo banco configurado em DATABASE_URL quando existir; caso contrário,
    usa o SQLite local de compatibilidade.
    """
    with auth_get_conn() as conn:
        db_execute(conn, f"""CREATE TABLE IF NOT EXISTS encaminhamento_empresas (
            id {db_id_type()},
            nome TEXT NOT NULL UNIQUE,
            criado_em TEXT NOT NULL
        )""")
        db_execute(conn, f"""CREATE TABLE IF NOT EXISTS encaminhamento_especialistas (
            id {db_id_type()},
            nome TEXT NOT NULL UNIQUE,
            criado_em TEXT NOT NULL
        )""")
        conn.commit()


def encaminhamento_listar_cadastros(table_name: str):
    if table_name not in {'encaminhamento_empresas', 'encaminhamento_especialistas'}:
        return []
    with auth_get_conn() as conn:
        return db_fetchall(conn, f"SELECT id, nome FROM {table_name} ORDER BY nome")


def encaminhamento_add_cadastro(table_name: str, nome: str) -> tuple[bool, str]:
    if table_name == 'encaminhamento_empresas':
        nome = encaminhamento_empresa_name(nome)
        label = 'Empresa'
    elif table_name == 'encaminhamento_especialistas':
        nome = encaminhamento_especialista_name(nome)
        label = 'Especialista'
    else:
        return False, 'Cadastro inválido.'
    if not nome:
        return False, f'Informe o nome de {label.lower()}.'
    try:
        with auth_get_conn() as conn:
            db_execute(conn, f"INSERT INTO {table_name} (nome, criado_em) VALUES ({db_param()}, {db_param()})", (nome, datetime.now().isoformat(timespec='seconds')))
            conn.commit()
        return True, f'{label} cadastrado(a) com sucesso.'
    except Exception:
        logger.exception('Erro ao cadastrar item de encaminhamento')
        return False, f'Não foi possível cadastrar. Talvez {label.lower()} já exista.'


def encaminhamento_delete_cadastro(table_name: str, item_id: str) -> tuple[bool, str]:
    if table_name not in {'encaminhamento_empresas', 'encaminhamento_especialistas'}:
        return False, 'Cadastro inválido.'
    try:
        with auth_get_conn() as conn:
            db_execute(conn, f"DELETE FROM {table_name} WHERE id = {db_param()}", (item_id,))
            conn.commit()
        return True, 'Cadastro removido com sucesso.'
    except Exception:
        logger.exception('Erro ao remover cadastro de encaminhamento')
        return False, 'Não foi possível remover o cadastro.'


def encaminhamento_format_date(raw_date: str) -> str:
    if raw_date:
        dt = datetime.strptime(raw_date, '%Y-%m-%d').date()
    else:
        dt = datetime.today().date()
    return dt.strftime('%d/%m/%Y')


def docx_replace_paragraph(paragraph, text: str, bold: bool = False, alignment=None):
    for run in list(paragraph.runs):
        run.text = ''
    run = paragraph.add_run(text)
    run.bold = bold
    if alignment is not None:
        paragraph.alignment = alignment


def docx_replace_paragraph_parts(paragraph, parts, alignment=None):
    """Substitui o parágrafo preservando a regra de negrito por trecho.

    parts: lista de tuplas (texto, bold).
    """
    for run in list(paragraph.runs):
        run.text = ''
    for text, bold in parts:
        if text:
            run = paragraph.add_run(text)
            run.bold = bool(bold)
    if alignment is not None:
        paragraph.alignment = alignment


def encaminhamento_extract_logo(tmpdir: str) -> str | None:
    """Extrai a logo do modelo original para reutilizar no DOCX gerado.

    A versão anterior editava/duplicava diretamente a estrutura do template. Isso
    herdava bordas e tabelas invisíveis do Word, causando a linha encostada no
    segundo encaminhamento. Agora a saída é montada do zero, usando apenas a logo.
    """
    try:
        with zipfile.ZipFile(ENCAMINHAMENTO_PREENCHIMENTO_TEMPLATE_PATH) as zf:
            media_files = [name for name in zf.namelist() if name.startswith('word/media/')]
            if not media_files:
                return None
            media_name = media_files[0]
            ext = os.path.splitext(media_name)[1] or '.jpeg'
            logo_path = os.path.join(tmpdir, f'encaminhamento_logo{ext}')
            with open(logo_path, 'wb') as out:
                out.write(zf.read(media_name))
            return logo_path
    except Exception:
        logger.exception('Não foi possível extrair a logo do template de encaminhamento')
        return None


def encaminhamento_set_cell_border(cell, color: str = '111827', size: str = '8'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in('w:tcBorders')
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right'):
        tag = 'w:{}'.format(edge)
        element = tcBorders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            tcBorders.append(element)
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), size)
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), color)


def encaminhamento_clear_cell_padding(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in('w:tcMar')
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar')
        tcPr.append(tcMar)
    for m in ('top', 'left', 'bottom', 'right'):
        node = tcMar.find(qn(f'w:{m}'))
        if node is None:
            node = OxmlElement(f'w:{m}')
            tcMar.append(node)
        node.set(qn('w:w'), '130')
        node.set(qn('w:type'), 'dxa')


def encaminhamento_add_paragraph(cell, parts=None, text: str = '', bold: bool = False, size: int = 10,
                                 align=None, space_after: int = 4, space_before: int = 0):
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = 1.0
    if align is not None:
        p.alignment = align
    if parts:
        for value, is_bold in parts:
            run = p.add_run(value)
            run.bold = bool(is_bold)
            run.font.size = Pt(size)
            run.font.name = 'Arial'
    else:
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = 'Arial'
    return p


def encaminhamento_add_block(doc, logo_path: str | None, empresa_upper: str, funcionario_upper: str,
                             funcionario_frase: str, rg_clean: str, cpf_clean: str,
                             especialista_clean: str, data_fmt: str, local_exame: dict):
    """Cria um encaminhamento limpo, sem copiar tabelas/linhas residuais do modelo."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(7.25)
    cell = table.cell(0, 0)
    cell.width = Inches(7.25)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    encaminhamento_set_cell_border(cell)
    encaminhamento_clear_cell_padding(cell)

    # Remove o parágrafo vazio padrão da célula, evitando espaços fantasmas.
    if cell.paragraphs:
        p = cell.paragraphs[0]._element
        p.getparent().remove(p)

    if logo_path and os.path.exists(logo_path):
        logo_p = cell.add_paragraph()
        logo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_p.paragraph_format.space_after = Pt(4)
        logo_run = logo_p.add_run()
        logo_run.add_picture(logo_path, width=Inches(5.25))

    encaminhamento_add_paragraph(cell, text='ENCAMINHAMENTO', bold=True, size=15,
                                 align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    encaminhamento_add_paragraph(cell, parts=[('Empresa : ', True), (empresa_upper, False)], size=9, space_after=3)
    encaminhamento_add_paragraph(cell, parts=[
        ('Funcionário: ', True), (funcionario_upper, False),
        (' RG: ', True), (rg_clean, False),
        (' CPF: ', True), (cpf_clean, False),
    ], size=9, space_after=12)

    encaminhamento_add_paragraph(
        cell,
        text=f'Encaminho o(a) colaborador(a) {funcionario_frase}, para avaliação com especialista {especialista_clean}.',
        bold=True,
        size=10,
        space_after=18,
    )

    encaminhamento_add_paragraph(cell, text=f"{local_exame['cidade_data']}, {data_fmt}", bold=True, size=10,
                                 align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=16)
    encaminhamento_add_paragraph(cell, text='__________________________', bold=True, size=10,
                                 align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=0)
    encaminhamento_add_paragraph(cell, text='MÉDICO EXAMINADOR', bold=True, size=10,
                                 align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=18)

    encaminhamento_add_paragraph(cell, text=f"Endereço: {local_exame['endereco']}",
                                 bold=True, size=8, space_after=0)
    encaminhamento_add_paragraph(cell, text=f"Fone: {local_exame['telefone']} e-mail: edgeocupacional@hotmail.com",
                                 bold=True, size=8, space_after=0)


def gerar_encaminhamento_especialista_docx(empresa: str, funcionario: str, rg: str, cpf: str, data_doc: str, especialista: str, local_exame_key: str, output_path: str):
    empresa_upper = encaminhamento_empresa_name(empresa)
    funcionario_upper = encaminhamento_clean_text(funcionario, upper=True)
    funcionario_frase = encaminhamento_title_name(funcionario)
    rg_clean = encaminhamento_clean_text(rg)
    cpf_clean = encaminhamento_clean_text(cpf)
    especialista_clean = encaminhamento_especialista_name(especialista)
    data_fmt = encaminhamento_format_date(data_doc)
    local = encaminhamento_get_local_exame(local_exame_key)
    if not local:
        raise ValueError('Local de exame inválido.')

    replacements = {
        '{{EMPRESA - CNPJ}}': empresa_upper,
        '{{NOME}}': funcionario_upper,
        '{{Nome}}': funcionario_frase,
        '{{RG}}': rg_clean,
        '{{CPF}}': cpf_clean,
        '{{tipo de especialista}}': especialista_clean,
        '{{CIDADE, UF}}': local['cidade_uf'],
        '{{DATA}}': data_fmt,
        '{{DEPENDE DA CIDADE}}': local['endereco'],
        '{{ENDEREÇO}}': local['endereco'],
        '{{telefone}}': local['telefone'],
    }
    replace_docx_placeholders_preserve_layout(ENCAMINHAMENTO_PREENCHIMENTO_TEMPLATE_PATH, output_path, replacements)


def encaminhamento_especialista_render(form_data=None):
    form_data = form_data or {}
    return render_template('encaminhamento_especialista.html',
                           title='Encaminhamento',
                           today=form_data.get('data_documento') or datetime.today().strftime('%Y-%m-%d'),
                           empresas=encaminhamento_listar_cadastros('encaminhamento_empresas'),
                           especialistas=encaminhamento_listar_cadastros('encaminhamento_especialistas'),
                           locais_exame=ENCAMINHAMENTO_LOCAIS_EXAME,
                           form_data=form_data)


init_encaminhamento_db()


@app.route('/encaminhamento-especialista', methods=['GET'])
def encaminhamento_especialista():
    return encaminhamento_especialista_render()


@app.route('/encaminhamento-especialista/cadastros', methods=['GET'])
def encaminhamento_especialista_cadastros():
    return render_template('encaminhamento_especialista_cadastros.html',
                           title='Cadastros de Encaminhamento',
                           empresas=encaminhamento_listar_cadastros('encaminhamento_empresas'),
                           especialistas=encaminhamento_listar_cadastros('encaminhamento_especialistas'))


@app.route('/encaminhamento-especialista/cadastros/empresa', methods=['POST'])
def encaminhamento_especialista_add_empresa():
    ok, msg = encaminhamento_add_cadastro('encaminhamento_empresas', request.form.get('nome', ''))
    flash(msg, 'success' if ok else 'error')
    return redirect(url_for('encaminhamento_especialista_cadastros'))


@app.route('/encaminhamento-especialista/cadastros/especialista', methods=['POST'])
def encaminhamento_especialista_add_especialista():
    ok, msg = encaminhamento_add_cadastro('encaminhamento_especialistas', request.form.get('nome', ''))
    flash(msg, 'success' if ok else 'error')
    return redirect(url_for('encaminhamento_especialista_cadastros'))


@app.route('/encaminhamento-especialista/cadastros/remover', methods=['POST'])
def encaminhamento_especialista_remover_cadastro():
    tipo = request.form.get('tipo')
    table = 'encaminhamento_empresas' if tipo == 'empresa' else 'encaminhamento_especialistas' if tipo == 'especialista' else ''
    ok, msg = encaminhamento_delete_cadastro(table, request.form.get('id', ''))
    flash(msg, 'success' if ok else 'error')
    return redirect(url_for('encaminhamento_especialista_cadastros'))


@app.route('/encaminhamento-especialista/gerar', methods=['POST'])
def encaminhamento_especialista_gerar():
    form_data = request.form.to_dict(flat=True)
    empresa = encaminhamento_empresa_name(request.form.get('empresa', ''))
    funcionario = encaminhamento_clean_text(request.form.get('funcionario', ''), upper=True)
    rg = encaminhamento_clean_text(request.form.get('rg', ''))
    cpf = encaminhamento_clean_text(request.form.get('cpf', ''))
    data_documento = request.form.get('data_documento', '')
    especialista = encaminhamento_especialista_name(request.form.get('especialista', ''))
    local_exame_key = request.form.get('local_exame', '')

    if not empresa or not funcionario or not rg or not cpf or not especialista or not local_exame_key:
        flash('Preencha empresa, funcionário, RG, CPF, especialista e local do exame.', 'error')
        return encaminhamento_especialista_render(form_data)

    if not encaminhamento_get_local_exame(local_exame_key):
        flash('Selecione um local de exame válido: Belém-PA ou Macapá-AP.', 'error')
        return encaminhamento_especialista_render(form_data)

    if len(somente_numeros(cpf)) != 11:
        flash('CPF inválido. Informe os 11 números do CPF.', 'error')
        return encaminhamento_especialista_render(form_data)

    filename_base = sanitize_filename(f'ENCAMINHAMENTO - {funcionario}')
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = os.path.join(tmpdir, f'{filename_base}.docx')
        try:
            gerar_encaminhamento_especialista_docx(empresa, funcionario, rg, cpf, data_documento, especialista, local_exame_key, output_path)
        except Exception:
            logger.exception('Erro ao gerar encaminhamento para especialista')
            flash('Não foi possível gerar o encaminhamento. Confira os dados e tente novamente.', 'error')
            return encaminhamento_especialista_render(form_data)
        payload = Path(output_path).read_bytes()
        return send_file(io.BytesIO(payload), as_attachment=True, download_name=f'{filename_base}.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


# =========================
# EXAMES A PRAZO
# =========================
def _exames_a_prazo_base_dir() -> str:
    path = os.path.join(DATA_DIR, "exames_a_prazo")
    os.makedirs(path, exist_ok=True)
    return path


def _exames_a_prazo_token_ok(token: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9_-]{12,80}", token or ""))


def _exames_a_prazo_job_dir(token: str) -> str:
    if not _exames_a_prazo_token_ok(token):
        raise ValueError("Sessão inválida. Reenvie as bases de consulta.")
    base = os.path.abspath(_exames_a_prazo_base_dir())
    path = os.path.abspath(os.path.join(base, token))
    if not (path == base or path.startswith(base + os.sep)):
        raise ValueError("Sessão inválida.")
    return path


def _read_upload_bytes(upload) -> bytes:
    try:
        upload.stream.seek(0)
    except Exception:
        pass
    raw = upload.read()
    try:
        upload.stream.seek(0)
    except Exception:
        pass
    return raw or b""


def _exames_a_prazo_cleanup_old(max_age_hours: int = 24) -> None:
    base = _exames_a_prazo_base_dir()
    now = datetime.now().timestamp()
    max_age = max_age_hours * 3600
    try:
        for name in os.listdir(base):
            full = os.path.join(base, name)
            if os.path.isdir(full) and now - os.path.getmtime(full) > max_age:
                shutil.rmtree(full, ignore_errors=True)
    except Exception as exc:
        logger.warning("Falha ao limpar sessões antigas de Exames a Prazo: %s", exc)


def _exames_a_prazo_load_context(token: str):
    job_dir = _exames_a_prazo_job_dir(token)
    manifest_path = os.path.join(job_dir, "manifest.json")
    if not os.path.exists(manifest_path):
        raise ValueError("As bases de consulta não foram encontradas. Reenvie as bases.")
    with open(manifest_path, "r", encoding="utf-8") as fh:
        manifest = json.load(fh)
    source_paths = [os.path.join(job_dir, name) for name in manifest.get("source_files", [])]
    source_files = []
    for path in source_paths:
        if os.path.exists(path):
            with open(path, "rb") as fh:
                source_files.append(fh.read())
    if not source_files:
        raise ValueError("Nenhuma base de consulta válida foi localizada. Reenvie as bases.")
    return manifest, source_files


def _exames_a_prazo_expand_request_uploads(uploads):
    extracted = []
    errors = []

    for upload in uploads or []:
        if not upload or not getattr(upload, "filename", ""):
            continue
        original_name = Path(upload.filename).name
        suffix = Path(original_name).suffix.lower()
        raw = _read_upload_bytes(upload)
        if not raw:
            errors.append(f"{original_name}: arquivo vazio.")
            continue

        if suffix == ".xlsx":
            if len(raw) > EXAMES_A_PRAZO_MAX_SINGLE_XLSX_BYTES:
                errors.append(f"{original_name}: arquivo maior que o limite de 30 MB.")
                continue
            extracted.append((original_name, raw))
            continue

        if suffix == ".zip":
            try:
                with zipfile.ZipFile(BytesIO(raw), "r") as zf:
                    infos = [
                        info for info in zf.infolist()
                        if not info.is_dir()
                        and Path(info.filename).suffix.lower() == ".xlsx"
                        and not Path(info.filename).name.startswith("~$")
                        and "__MACOSX" not in Path(info.filename).parts
                    ]
                    if not infos:
                        errors.append(f"{original_name}: o ZIP não contém nenhuma planilha .xlsx.")
                        continue
                    if len(infos) > EXAMES_A_PRAZO_MAX_REQUEST_FILES:
                        errors.append(f"{original_name}: o ZIP contém {len(infos)} planilhas; o limite é {EXAMES_A_PRAZO_MAX_REQUEST_FILES}.")
                        continue
                    if sum(info.file_size for info in infos) > EXAMES_A_PRAZO_MAX_ZIP_UNCOMPRESSED_BYTES:
                        errors.append(f"{original_name}: o conteúdo descompactado ultrapassa 250 MB.")
                        continue

                    for info in infos:
                        filename = Path(info.filename).name
                        if info.file_size > EXAMES_A_PRAZO_MAX_SINGLE_XLSX_BYTES:
                            errors.append(f"{filename}: arquivo maior que o limite de 30 MB e foi ignorado.")
                            continue
                        extracted.append((filename, zf.read(info)))
            except zipfile.BadZipFile:
                errors.append(f"{original_name}: ZIP inválido ou corrompido.")
            continue

        errors.append(f"{original_name}: formato não aceito. Envie .xlsx ou .zip.")

    unique = []
    seen = {}
    for filename, content in extracted:
        key = filename.casefold()
        if key in seen:
            errors.append(
                f"Nome duplicado '{filename}'. Como a saída precisa manter exatamente o nome original, deixe apenas uma planilha com esse nome."
            )
            continue
        seen[key] = filename
        unique.append((filename, content))

    if len(unique) > EXAMES_A_PRAZO_MAX_REQUEST_FILES:
        errors.append(f"Foram recebidas mais de {EXAMES_A_PRAZO_MAX_REQUEST_FILES} planilhas; apenas as primeiras foram consideradas.")
        unique = unique[:EXAMES_A_PRAZO_MAX_REQUEST_FILES]

    return unique, errors


@app.route('/exames-a-prazo', methods=['GET'])
def exames_a_prazo():
    return render_template(
        'exames_a_prazo.html',
        title='Exames a Prazo',
        month_options=[],
        source_token='',
        source_names=[],
        selected_months=[],
        stage='upload',
        max_upload_mb=get_max_upload_mb(),
    )


@app.route('/exames-a-prazo/guias', methods=['POST'])
def exames_a_prazo_guias():
    from edge_app.exames_a_prazo_core import get_month_sheets_from_sources

    _exames_a_prazo_cleanup_old()
    uploads = [request.files.get('source_base_1'), request.files.get('source_base_2')]
    source_files = []
    source_names = []
    messages = []

    for idx, upload in enumerate(uploads, 1):
        if not upload or not getattr(upload, 'filename', ''):
            continue
        ok, msg = validate_uploaded_file(upload, EXAMES_A_PRAZO_SOURCE_EXTENSIONS, f'a base de consulta {idx}')
        if not ok:
            messages.append(msg)
            continue
        raw = _read_upload_bytes(upload)
        if not raw:
            messages.append(f"{Path(upload.filename).name}: arquivo vazio.")
            continue
        if len(raw) > EXAMES_A_PRAZO_MAX_SINGLE_XLSX_BYTES:
            messages.append(f"{Path(upload.filename).name}: arquivo maior que o limite configurado de {EXAMES_A_PRAZO_MAX_SINGLE_XLSX_BYTES // (1024 * 1024)} MB.")
            continue
        source_files.append(raw)
        source_names.append(Path(upload.filename).name)

    if messages:
        for msg in messages:
            flash(msg, 'error')
    if not source_files:
        flash('Envie pelo menos uma base de consulta em .xlsx.', 'error')
        return redirect(url_for('exames_a_prazo'))

    try:
        month_options = get_month_sheets_from_sources(source_files)
    except Exception as exc:
        logger.exception('Erro ao ler bases de consulta do Exames a Prazo')
        flash(f'Não foi possível ler as bases de consulta: {exc}', 'error')
        return redirect(url_for('exames_a_prazo'))

    if not month_options:
        flash('Nenhuma guia mensal foi identificada nas bases enviadas.', 'error')
        return redirect(url_for('exames_a_prazo'))

    token = secrets.token_urlsafe(24)
    job_dir = _exames_a_prazo_job_dir(token)
    os.makedirs(job_dir, exist_ok=True)
    saved = []
    for idx, raw in enumerate(source_files, 1):
        filename = f'source_{idx}.xlsx'
        with open(os.path.join(job_dir, filename), 'wb') as fh:
            fh.write(raw)
        saved.append(filename)

    manifest = {
        'created_at': datetime.now().isoformat(timespec='seconds'),
        'source_files': saved,
        'source_names': source_names,
        'month_options': month_options,
    }
    with open(os.path.join(job_dir, 'manifest.json'), 'w', encoding='utf-8') as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)

    default_months = month_options[-2:] if len(month_options) >= 2 else month_options
    return render_template(
        'exames_a_prazo.html',
        title='Exames a Prazo',
        month_options=month_options,
        source_token=token,
        source_names=source_names,
        selected_months=default_months,
        stage='generate',
        max_upload_mb=get_max_upload_mb(),
    )


@app.route('/exames-a-prazo/gerar', methods=['POST'])
def exames_a_prazo_gerar():
    from edge_app.exames_a_prazo_core import (
        company_key,
        extract_cnpjs_from_workbook,
        format_cnpj,
        generate_group_workbook,
        generate_solo_workbook,
        month_identity,
        normalize_text,
        parse_selected_months_from_sources,
    )

    token = request.form.get('source_token', '').strip()
    selected_months = [m.strip() for m in request.form.getlist('selected_months') if m.strip()]
    try:
        manifest, source_files = _exames_a_prazo_load_context(token)
    except Exception as exc:
        flash(str(exc), 'error')
        return redirect(url_for('exames_a_prazo'))

    if not selected_months:
        flash('Selecione pelo menos uma guia/competência.', 'error')
        return render_template(
            'exames_a_prazo.html',
            title='Exames a Prazo',
            month_options=manifest.get('month_options', []),
            source_token=token,
            source_names=manifest.get('source_names', []),
            selected_months=[],
            stage='generate',
            max_upload_mb=get_max_upload_mb(),
        )

    # Validação tolerante: algumas abas do Excel podem ter espaço oculto no
    # final do nome (ex.: "JANEIRO.2025 OK "). O checkbox da página pode voltar
    # como "JANEIRO.2025 OK", então não podemos comparar apenas texto exato.
    month_options = [str(m).strip() for m in manifest.get('month_options', []) if str(m).strip()]
    option_exact = set(month_options)
    option_norms = {normalize_text(m) for m in month_options}
    option_ids = {month_identity(m) for m in month_options if month_identity(m) is not None}

    invalid_months = []
    for month in selected_months:
        month_id = month_identity(month)
        if month in option_exact or normalize_text(month) in option_norms or (month_id is not None and month_id in option_ids):
            continue
        invalid_months.append(month)

    if invalid_months:
        flash(
            'Uma ou mais guias selecionadas não pertencem às bases carregadas. Recarregue as bases e tente novamente. '
            f'Guias não reconhecidas: {", ".join(invalid_months)}',
            'error'
        )
        return redirect(url_for('exames_a_prazo'))

    request_uploads = request.files.getlist('request_files')
    request_files, expansion_errors = _exames_a_prazo_expand_request_uploads(request_uploads)
    errors = list(expansion_errors)
    requests_data = []
    all_requested_keys = []
    seen_requested_keys = set()

    for filename, content in request_files:
        try:
            cnpjs = extract_cnpjs_from_workbook(content)
        except Exception as exc:
            errors.append(f'{filename}: não foi possível ler a planilha ({exc}).')
            continue
        if not cnpjs:
            errors.append(f'{filename}: nenhum CNPJ foi encontrado e a planilha foi ignorada.')
            continue
        requests_data.append({'name': filename, 'bytes': content, 'cnpjs': cnpjs})
        for cnpj in cnpjs:
            key = company_key(cnpj)
            if key not in seen_requested_keys:
                seen_requested_keys.add(key)
                all_requested_keys.append(key)

    if not requests_data:
        for msg in errors or ['Envie pelo menos uma planilha .xlsx ou ZIP contendo CNPJs.']:
            flash(msg, 'error')
        return render_template(
            'exames_a_prazo.html',
            title='Exames a Prazo',
            month_options=manifest.get('month_options', []),
            source_token=token,
            source_names=manifest.get('source_names', []),
            selected_months=selected_months,
            stage='generate',
            max_upload_mb=get_max_upload_mb(),
        )

    if not os.path.exists(EXAMES_A_PRAZO_SOLO_TEMPLATE_PATH) or not os.path.exists(EXAMES_A_PRAZO_GROUP_TEMPLATE_PATH):
        flash('Os modelos do Exames a Prazo não foram encontrados no sistema.', 'error')
        return redirect(url_for('exames_a_prazo'))

    try:
        records, companies = parse_selected_months_from_sources(
            source_files,
            selected_months,
            all_requested_keys,
        )
    except Exception as exc:
        logger.exception('Erro ao analisar guias do Exames a Prazo')
        flash(f'Erro ao analisar as guias selecionadas: {exc}', 'error')
        return render_template(
            'exames_a_prazo.html',
            title='Exames a Prazo',
            month_options=manifest.get('month_options', []),
            source_token=token,
            source_names=manifest.get('source_names', []),
            selected_months=selected_months,
            stage='generate',
            max_upload_mb=get_max_upload_mb(),
        )

    output_files = {}
    for req in requests_data:
        filename = req['name']
        cnpjs = req['cnpjs']
        keys = [company_key(cnpj) for cnpj in cnpjs]
        request_companies = {
            key: companies.get(key, format_cnpj(cnpj))
            for key, cnpj in zip(keys, cnpjs)
        }
        try:
            if len(keys) == 1:
                key = keys[0]
                content = generate_solo_workbook(
                    records,
                    selected_months,
                    key,
                    request_companies[key],
                    EXAMES_A_PRAZO_SOLO_TEMPLATE_PATH,
                    include_empty_months=True,
                    empty_title_only=True,
                )
            else:
                content = generate_group_workbook(
                    records,
                    selected_months,
                    keys,
                    request_companies,
                    EXAMES_A_PRAZO_GROUP_TEMPLATE_PATH,
                    include_empty_months=True,
                    empty_title_only=True,
                )
            if content:
                output_files[filename] = content
            else:
                errors.append(f'{filename}: não foi possível montar a planilha.')
        except Exception as exc:
            logger.exception('Erro ao gerar planilha Exames a Prazo')
            errors.append(f'{filename}: {exc}')

    if not output_files:
        for msg in errors or ['Nenhum arquivo foi gerado.']:
            flash(msg, 'error')
        return render_template(
            'exames_a_prazo.html',
            title='Exames a Prazo',
            month_options=manifest.get('month_options', []),
            source_token=token,
            source_names=manifest.get('source_names', []),
            selected_months=selected_months,
            stage='generate',
            max_upload_mb=get_max_upload_mb(),
        )

    # Relatório técnico simples dentro do ZIP para o usuário conferir rapidamente
    # se a base foi lida e quantos registros foram encontrados por CNPJ.
    record_counts = {key: 0 for key in all_requested_keys}
    for rec in records:
        if rec.company_key in record_counts:
            record_counts[rec.company_key] += 1
    resumo_lines = [
        'RESUMO DA GERAÇÃO - EXAMES A PRAZO',
        '',
        f'Competências selecionadas: {", ".join(selected_months)}',
        f'Total de CNPJs solicitados: {len(all_requested_keys)}',
        f'Total de registros encontrados: {len(records)}',
        '',
        'CNPJs / empresas:',
    ]
    for req in requests_data:
        for cnpj in req['cnpjs']:
            key = company_key(cnpj)
            resumo_lines.append(f'- {format_cnpj(cnpj)}: {record_counts.get(key, 0)} registro(s)')
    if errors:
        resumo_lines.extend(['', 'Avisos:', *errors])

    bio = BytesIO()
    with zipfile.ZipFile(bio, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for filename, content in output_files.items():
            zf.writestr(sanitize_filename(filename), content)
        zf.writestr('RESUMO_EXAMES_A_PRAZO.txt', '\n'.join(resumo_lines))
        if errors:
            zf.writestr('ATENCAO_ERROS.txt', '\n'.join(errors))
    bio.seek(0)

    audit_log('exames_a_prazo_gerado', f'{len(output_files)} planilha(s); guias: {", ".join(selected_months)}')
    return send_file(
        bio,
        mimetype='application/zip',
        as_attachment=True,
        download_name='EXAMES_A_PRAZO.zip',
    )

@app.route("/healthz")
def healthz():
    return jsonify({"ok": True, "app": APP_TITLE, "time": datetime.utcnow().isoformat() + "Z"})


@app.route("/status-sistema")
def status_sistema():
    class Card:
        def __init__(self, label, value, detail, ok=True):
            self.label = label
            self.value = value
            self.detail = detail
            self.ok = ok

    data_dir = DATA_DIR
    envio_dir = os.environ.get("ENVIO_PERIODICOS_DATA_DIR") or os.path.join(os.environ.get("RENDER_DISK_PATH", DATA_DIR), "envio_periodicos")
    separar_dir = os.environ.get("SEPARAR_EXAMES_DATA_DIR") or os.path.join(os.environ.get("RENDER_DISK_PATH", DATA_DIR), "separar_exames")

    db_ok = False
    db_detail = ""
    try:
        if USE_POSTGRES and psycopg2:
            conn = psycopg2.connect(DATABASE_URL, connect_timeout=5)
            cur = conn.cursor(); cur.execute("SELECT 1"); cur.close(); conn.close()
            db_ok = True; db_detail = "PostgreSQL conectado"
        else:
            os.makedirs(data_dir, exist_ok=True)
            db_ok = True; db_detail = "SQLite/local disponível"
    except Exception as exc:
        db_detail = f"Falha: {exc}"

    tesseract = shutil.which("tesseract") or ""
    ocr_ok = bool(tesseract)
    disk_ok = False
    disk_text = "Não foi possível medir o disco"
    try:
        os.makedirs(data_dir, exist_ok=True)
        usage = shutil.disk_usage(data_dir)
        total_gb = usage.total / (1024**3)
        free_gb = usage.free / (1024**3)
        disk_text = f"{free_gb:.1f} GB livres de {total_gb:.1f} GB"
        disk_ok = free_gb > 0.2
    except Exception:
        pass

    persist_root = os.environ.get("RENDER_DISK_PATH") or ""
    persist_ok = bool(persist_root and os.path.isdir(persist_root))
    cards = [
        Card("Banco de dados", "Conectado" if db_ok else "Verificar", db_detail, db_ok),
        Card("Disco persistente", "Ativo" if persist_ok else "Não confirmado", persist_root or "RENDER_DISK_PATH não configurado", persist_ok),
        Card("OCR / Tesseract", "Ativo" if ocr_ok else "Não localizado", tesseract or "PDFs escaneados podem não ser lidos", ocr_ok),
        Card("Espaço em disco", "Disponível" if disk_ok else "Verificar", disk_text, disk_ok),
        Card("Envio periódicos", "Pasta configurada" if envio_dir else "Verificar", envio_dir, bool(envio_dir)),
        Card("Separar exames", "Pasta configurada" if separar_dir else "Verificar", separar_dir, bool(separar_dir)),
    ]
    return render_template(
        "status_sistema.html",
        title="Status do sistema",
        cards=cards,
        data_dir=data_dir,
        envio_dir=envio_dir,
        separar_dir=separar_dir,
        tesseract=tesseract,
        disk_text=disk_text,
        env_name=os.environ.get("FLASK_ENV") or os.environ.get("RENDER_SERVICE_NAME") or "produção",
    )

@app.route("/")
def home():
    return render_template("home.html", title=APP_TITLE)

@app.route("/relatorios", methods=["GET", "POST"])
def relatorios():
    if request.method == "POST":
        files = request.files.getlist("files")
        try:
            mes = int(request.form.get("mes"))
            if mes < 1 or mes > 12:
                raise ValueError
        except Exception:
            flash("Mês inválido.")
            return redirect(url_for("relatorios"))

        try:
            arquivos_memoria = extrair_planilhas_relatorios_uploads(files)
        except ValueError as exc:
            flash(str(exc))
            return redirect(url_for("relatorios"))

        if not arquivos_memoria:
            flash("Selecione planilhas em .xls/.xlsx ou um ZIP contendo as planilhas.")
            return redirect(url_for("relatorios"))

        files_rel = [UploadedMemoryFile(nome, dados) for nome, dados in arquivos_memoria]
        files_base = [UploadedMemoryFile(nome, dados) for nome, dados in arquivos_memoria]

        try:
            cadastros = carregar_empresas_cnpj_relatorios()
            wb_rel = criar_relatorio(files_rel, mes, cadastros)
            wb_base = criar_base(files_base, mes, cadastros)
        except Exception:
            logger.exception("Erro ao gerar relatórios")
            flash("Não foi possível gerar os relatórios. Confira se as planilhas estão no modelo esperado.")
            return redirect(url_for("relatorios"))

        temp_dir = tempfile.mkdtemp()
        caminho_rel = os.path.join(temp_dir, f"Relatorio_{nome_mes(mes)}.xlsx")
        caminho_base = os.path.join(temp_dir, f"Base_do_Mes_{nome_mes(mes)}.xlsx")
        caminho_zip = os.path.join(temp_dir, f"Arquivos_{nome_mes(mes)}.zip")
        wb_rel.save(caminho_rel)
        wb_base.save(caminho_base)

        with zipfile.ZipFile(caminho_zip, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(caminho_rel, os.path.basename(caminho_rel))
            z.write(caminho_base, os.path.basename(caminho_base))

        return send_file(caminho_zip, as_attachment=True, download_name=f"Arquivos_{nome_mes(mes)}.zip")
    return render_template("relatorios.html", empresas_cnpj=listar_empresas_cnpj_relatorios())


@app.route("/relatorios/empresas-cnpj/modelo")
def relatorios_empresas_cnpj_modelo():
    wb = Workbook()
    ws = wb.active
    ws.title = "Empresas CNPJ"
    headers = ["EMPRESA", "CNPJ"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    ws.cell(row=2, column=1, value="EMPRESA EXEMPLO LTDA")
    ws.cell(row=2, column=2, value="12.345.678/0001-90")
    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 22
    temp_dir = tempfile.mkdtemp(prefix="modelo_empresas_cnpj_")
    path = os.path.join(temp_dir, "MODELO_CADASTRO_EMPRESAS_CNPJ.xlsx")
    wb.save(path)
    return send_file(path, as_attachment=True, download_name="MODELO_CADASTRO_EMPRESAS_CNPJ.xlsx")


@app.route("/relatorios/empresas-cnpj/importar", methods=["POST"])
def relatorios_empresas_cnpj_importar():
    try:
        total = importar_empresas_cnpj_relatorios(request.files.get("empresas_cnpj_file"))
        flash(f"Cadastro atualizado: {total} empresa(s)/CNPJ(s) importado(s) ou atualizado(s).")
    except Exception as exc:
        logger.exception("Erro ao importar cadastro de empresas/CNPJ")
        flash(str(exc) or "Não foi possível importar o cadastro de empresas/CNPJ.")
    return redirect(url_for("relatorios"))


# =========================
# ENCAMINHAMENTO DE EXAMES COMPLEMENTARES
# =========================
ENCAMINHAMENTO_COMPLEMENTARES_LOCAIS = CLINIC_LOCATIONS


def encaminhamento_complementares_get_local(value: str) -> dict | None:
    return ENCAMINHAMENTO_COMPLEMENTARES_LOCAIS.get((value or '').strip().lower())


def encaminhamento_complementares_render(form_data=None, exames=None):
    if form_data is None:
        form_data = {}
    if exames is None:
        exames = ['']
    return render_template(
        'encaminhamento_complementares.html',
        title='Encaminhamento de Exames Complementares',
        form_data=form_data,
        exames=exames,
        locais_exame=ENCAMINHAMENTO_COMPLEMENTARES_LOCAIS,
        today=datetime.today().strftime('%Y-%m-%d'),
    )


def encaminhamento_complementares_exam_values(raw_values) -> list[str]:
    exames = []
    for value in raw_values:
        clean = encaminhamento_clean_text(value, upper=True)
        if clean:
            exames.append(clean)
    return exames


def encaminhamento_complementares_iter_tables(container):
    for table in getattr(container, 'tables', []):
        yield table
        for row in table.rows:
            for cell in row.cells:
                yield from encaminhamento_complementares_iter_tables(cell)


def encaminhamento_complementares_iter_paragraphs(container):
    for paragraph in getattr(container, 'paragraphs', []):
        yield paragraph
    for table in getattr(container, 'tables', []):
        for row in table.rows:
            for cell in row.cells:
                yield from encaminhamento_complementares_iter_paragraphs(cell)


def encaminhamento_complementares_replace_runs(paragraph, replacements: dict[str, str]):
    for run in paragraph.runs:
        if not run.text:
            continue
        new_text = run.text
        for placeholder, value in replacements.items():
            new_text = new_text.replace(placeholder, value)
        run.text = new_text

    # Fallback para placeholders quebrados em vários runs pelo Word.
    full_text = paragraph.text
    if any(placeholder in full_text for placeholder in replacements):
        for placeholder, value in replacements.items():
            full_text = full_text.replace(placeholder, value)
        if paragraph.runs:
            first = paragraph.runs[0]
            for run in paragraph.runs[1:]:
                run.text = ''
            first.text = full_text
        else:
            paragraph.add_run(full_text)


def encaminhamento_complementares_set_cell_text(cell, text: str):
    if not cell.paragraphs:
        paragraph = cell.add_paragraph()
    else:
        paragraph = cell.paragraphs[0]
    for run in paragraph.runs:
        run.text = ''
    if paragraph.runs:
        run = paragraph.runs[0]
    else:
        run = paragraph.add_run()
    run.text = text
    run.bold = True
    run.font.name = 'Arial'
    run.font.size = Pt(10)


def encaminhamento_complementares_is_exam_table(table) -> bool:
    for row in table.rows:
        for cell in row.cells:
            if '{{EXAME' in cell.text:
                return True
    return False


def encaminhamento_complementares_apply_exam_rows(table, exames: list[str]):
    if not exames:
        return
    # Se houver mais exames do que linhas no modelo, cria linhas novas copiando o padrão da última linha.
    while len(table.rows) < len(exames):
        table._tbl.append(deepcopy(table.rows[-1]._tr))
    # Se houver menos exames, remove as linhas excedentes para que a tabela tenha exatamente a quantidade informada.
    while len(table.rows) > len(exames):
        table._tbl.remove(table.rows[-1]._tr)
    for idx, exame in enumerate(exames):
        encaminhamento_complementares_set_cell_text(table.rows[idx].cells[0], exame)


def encaminhamento_complementares_set_texts_in_element(element, placeholder: str, value: str):
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    for text_node in element.xpath('.//w:t', namespaces=ns):
        if text_node.text and placeholder in text_node.text:
            text_node.text = text_node.text.replace(placeholder, value)


def encaminhamento_complementares_replace_xml_placeholders(root, replacements: dict[str, str]) -> None:
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    for text_node in root.xpath('.//w:t', namespaces=ns):
        if not text_node.text:
            continue
        for placeholder, value in replacements.items():
            text_node.text = text_node.text.replace(placeholder, value)


def encaminhamento_complementares_apply_exam_rows_xml(root, exames: list[str]) -> None:
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    rows = list(root.xpath('.//w:tr[.//w:t[contains(., "{{EXAME")]]', namespaces=ns))
    for row in rows:
        # Processa somente as linhas internas da tabela de exames. As linhas externas
        # também enxergam o texto da tabela aninhada, por isso são ignoradas aqui.
        if row.xpath('.//w:tr', namespaces=ns):
            continue
        row_text = ''.join(text_node.text or '' for text_node in row.xpath('.//w:t', namespaces=ns))
        match = re.search(r'\{\{EXAME(\d+)\}\}', row_text)
        if not match:
            continue
        exam_number = int(match.group(1))
        parent = row.getparent()
        placeholder = f'{{{{EXAME{exam_number}}}}}'
        exam_index = exam_number - 1

        if exam_number == 5 and len(exames) > 5:
            insert_at = parent.index(row)
            parent.remove(row)
            for offset, exame in enumerate(exames[4:]):
                cloned = deepcopy(row)
                encaminhamento_complementares_set_texts_in_element(cloned, placeholder, exame)
                parent.insert(insert_at + offset, cloned)
        elif exam_index < len(exames):
            encaminhamento_complementares_set_texts_in_element(row, placeholder, exames[exam_index])
        else:
            parent.remove(row)

def gerar_encaminhamento_complementares_docx(nome: str, cpf: str, data_doc: str, exames: list[str], local_exame_key: str, output_path: str):
    local = encaminhamento_complementares_get_local(local_exame_key)
    if not local:
        raise ValueError('Local do exame inválido.')
    if not exames:
        raise ValueError('Informe pelo menos um exame.')

    replacements = {
        '{{NOME}}': encaminhamento_clean_text(nome, upper=True),
        '{{CPF}}': encaminhamento_clean_text(cpf),
        '{{CIDADE, UF}}': local['cidade_uf'],
        '{{DATA}}': encaminhamento_format_date(data_doc),
        '{{ENDEREÇO}}': local['endereco'],
        '{{telefone}}': local['telefone'],
    }

    parser = etree.XMLParser(remove_blank_text=False, recover=False)
    with zipfile.ZipFile(ENCAMINHAMENTO_COMPLEMENTARES_TEMPLATE_PATH, 'r') as zin, zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml':
                root = etree.fromstring(data, parser=parser)
                encaminhamento_complementares_apply_exam_rows_xml(root, exames)
                encaminhamento_complementares_replace_xml_placeholders(root, replacements)
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')
            zout.writestr(item, data)


@app.route('/encaminhamento-complementares', methods=['GET'])
def encaminhamento_complementares():
    return encaminhamento_complementares_render()


@app.route('/encaminhamento-complementares/gerar', methods=['POST'])
def encaminhamento_complementares_gerar():
    nome = encaminhamento_clean_text(request.form.get('nome', ''), upper=True)
    cpf = encaminhamento_clean_text(request.form.get('cpf', ''))
    data_documento = request.form.get('data_documento', '')
    local_exame_key = request.form.get('local_exame', '')
    exames_raw = request.form.getlist('exames')
    exames = encaminhamento_complementares_exam_values(exames_raw)
    form_data = request.form.to_dict()

    if not nome or not cpf or not data_documento or not local_exame_key:
        flash('Preencha nome, CPF, data e local do exame.', 'error')
        return encaminhamento_complementares_render(form_data, exames_raw or [''])
    if not encaminhamento_complementares_get_local(local_exame_key):
        flash('Selecione um local do exame válido.', 'error')
        return encaminhamento_complementares_render(form_data, exames_raw or [''])
    if not exames:
        flash('Informe pelo menos um exame.', 'error')
        return encaminhamento_complementares_render(form_data, exames_raw or [''])
    if len(exames) > 5:
        flash('Informe no máximo 5 exames por encaminhamento.', 'error')
        return encaminhamento_complementares_render(form_data, exames_raw or [''])
    try:
        datetime.strptime(data_documento, '%Y-%m-%d')
    except ValueError:
        flash('Data inválida.', 'error')
        return encaminhamento_complementares_render(form_data, exames_raw or [''])

    filename_base = sanitize_filename(f'ENCAMINHAMENTO COMPLEMENTARES - {nome}')
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = os.path.join(tmpdir, f'{filename_base}.docx')
        try:
            gerar_encaminhamento_complementares_docx(nome, cpf, data_documento, exames, local_exame_key, output_path)
        except Exception:
            logger.exception('Erro ao gerar encaminhamento de exames complementares')
            flash('Não foi possível gerar o encaminhamento. Confira os dados e tente novamente.', 'error')
            return encaminhamento_complementares_render(form_data, exames_raw or [''])
        payload = Path(output_path).read_bytes()
        return send_file(
            io.BytesIO(payload),
            as_attachment=True,
            download_name=f'{filename_base}.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )

@app.route("/encaminhamentos", methods=["GET", "POST"])
def encaminhamentos():
    flash("A geração de encaminhamentos agora fica dentro do módulo Envio periódicos.", "warning")
    return redirect("/envio-periodicos/")

@app.route("/renumerador", methods=["GET", "POST"])
def renumerador():
    if request.method == "POST":
        arquivo = request.files.get("arquivo")
        nova_data = request.form.get("nova_data", "").strip()

        if not arquivo or not arquivo.filename:
            flash("Selecione um arquivo.")
            return redirect(url_for("renumerador"))

        if not nova_data:
            flash("Informe a data.")
            return redirect(url_for("renumerador"))

        if not allowed_renum_file(arquivo.filename):
            flash("Formato inválido. Envie um arquivo .docx ou .zip.")
            return redirect(url_for("renumerador"))

        with tempfile.TemporaryDirectory() as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            entrada_dir = temp_dir / "entrada"
            saida_dir = temp_dir / "saida"
            entrada_dir.mkdir(parents=True, exist_ok=True)
            saida_dir.mkdir(parents=True, exist_ok=True)

            nome_seguro = secure_filename(arquivo.filename)
            caminho_upload = temp_dir / nome_seguro
            arquivo.save(caminho_upload)

            if caminho_upload.suffix.lower() == ".zip":
                try:
                    with zipfile.ZipFile(caminho_upload, "r") as zip_ref:
                        safe_extract_zip(caminho_upload, entrada_dir)
                except (zipfile.BadZipFile, ValueError) as exc:
                    logger.warning("ZIP inválido bloqueado no renumerador: %s", exc)
                    flash("O arquivo ZIP enviado está corrompido, inválido ou contém caminhos inseguros.")
                    return redirect(url_for("renumerador"))
            else:
                destino = entrada_dir / nome_seguro
                destino.write_bytes(caminho_upload.read_bytes())

            arquivos_docx = [p for p in entrada_dir.rglob("*.docx") if not p.name.startswith("~$")]
            if not arquivos_docx:
                flash("Nenhum arquivo .docx foi encontrado para processar.")
                return redirect(url_for("renumerador"))

            total_arquivos = 0
            total_recibos = 0
            relatorio = []

            for caminho in arquivos_docx:
                relativo = caminho.relative_to(entrada_dir)
                destino = saida_dir / relativo
                try:
                    resultado = renumerar_documento(str(caminho), str(destino), nova_data)
                    if len(resultado) == 3:
                        alterados, ultimo, datas_alteradas = resultado
                        avisos = []
                    else:
                        alterados, ultimo, datas_alteradas, avisos = resultado
                    total_arquivos += 1
                    total_recibos += alterados
                    detalhe_avisos = f" | avisos: {'; '.join(avisos)}" if avisos else ""
                    relatorio.append(f"{relativo.as_posix()} | último encontrado: {ultimo} | recibos renumerados: {alterados} | datas alteradas: {datas_alteradas}{detalhe_avisos}")
                except Exception as e:
                    relatorio.append(f"{relativo.as_posix()} | erro: {str(e)}")

            relatorio_path = saida_dir / "relatorio_processamento.txt"
            relatorio_path.write_text(
                "Renumerador de Recibos - Relatório de Processamento\n\n"
                f"Arquivos processados: {total_arquivos}\n"
                f"Recibos renumerados: {total_recibos}\n\n" + "\n".join(relatorio),
                encoding="utf-8"
            )

            if len(arquivos_docx) == 1:
                unico_saida = next((p for p in saida_dir.rglob("*.docx")), None)
                if unico_saida:
                    buffer = io.BytesIO(unico_saida.read_bytes())
                    buffer.seek(0)
                    return send_file(buffer, as_attachment=True, download_name=unico_saida.name, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            zip_path = temp_dir / "recibos_renumerados.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_out:
                for arquivo_saida in saida_dir.rglob("*"):
                    if arquivo_saida.is_file():
                        zip_out.write(arquivo_saida, arquivo_saida.relative_to(saida_dir))

            zip_buffer = io.BytesIO(zip_path.read_bytes())
            zip_buffer.seek(0)
            return send_file(zip_buffer, as_attachment=True, download_name="recibos_renumerados.zip", mimetype="application/zip")

    return render_template("renumerador.html")

def _get_esocial_base_session_state():
    """Retorna a planilha base persistida na sessão, se ainda existir no disco."""
    data = session.get("esocial_base")
    if not isinstance(data, dict):
        return None
    token = re.sub(r"[^A-Za-z0-9_-]", "", str(data.get("token", "")))
    stored_name = secure_filename(str(data.get("stored_name", "")))
    if not token or not stored_name:
        session.pop("esocial_base", None)
        return None
    path = Path(ESOCIAL_BASE_CACHE_DIR) / token / stored_name
    if not path.is_file():
        session.pop("esocial_base", None)
        return None
    sheets = [str(x) for x in (data.get("sheets") or []) if str(x).strip()]
    return {
        "token": token,
        "path": path,
        "filename": str(data.get("filename") or stored_name),
        "stored_name": stored_name,
        "sheets": sheets,
        "size": int(data.get("size") or path.stat().st_size),
    }


def _clear_esocial_base_session():
    data = session.pop("esocial_base", None)
    if isinstance(data, dict):
        token = re.sub(r"[^A-Za-z0-9_-]", "", str(data.get("token", "")))
        if token:
            shutil.rmtree(Path(ESOCIAL_BASE_CACHE_DIR) / token, ignore_errors=True)
    session.modified = True


def _persist_esocial_base_file(base_file):
    ok, msg = validate_uploaded_file(base_file, ESOCIAL_ALLOWED_EXTENSIONS, "a planilha base")
    if not ok:
        raise ValueError(msg)

    original_name = Path(base_file.filename).name
    suffix = Path(original_name).suffix.lower()
    token = secrets.token_urlsafe(18).replace("-", "_")
    folder = Path(ESOCIAL_BASE_CACHE_DIR) / token
    folder.mkdir(parents=True, exist_ok=True)
    stored_name = f"base{suffix}"
    path = folder / stored_name
    try:
        base_file.save(path)
        sheets = list_sheets(str(path))
        if not sheets:
            raise ValueError("A planilha base não possui guias disponíveis.")
    except Exception:
        shutil.rmtree(folder, ignore_errors=True)
        raise

    _clear_esocial_base_session()
    session["esocial_base"] = {
        "token": token,
        "filename": original_name,
        "stored_name": stored_name,
        "sheets": sheets,
        "size": path.stat().st_size,
    }
    session.modified = True
    return _get_esocial_base_session_state()


@app.route("/esocial", methods=["GET"])
def esocial():
    base_state = _get_esocial_base_session_state()
    base_info = None
    if base_state:
        base_info = {
            "filename": base_state["filename"],
            "sheets": base_state["sheets"],
            "size": base_state["size"],
        }
    return render_template("esocial.html", title="Recibo eSocial", esocial_base=base_info)


@app.route("/esocial/abas-base", methods=["POST"])
def esocial_abas_base():
    base_file = request.files.get("base_file")
    if not base_file or not base_file.filename:
        return jsonify({"ok": False, "error": "Nenhuma planilha base enviada."}), 400
    try:
        state = _persist_esocial_base_file(base_file)
        return jsonify({
            "ok": True,
            "filename": state["filename"],
            "size": state["size"],
            "sheets": state["sheets"],
        })
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception:
        logger.exception("Erro ao salvar/ler a planilha base do Recibo eSocial")
        return jsonify({"ok": False, "error": "Não foi possível ler as guias da planilha base."}), 500


@app.route("/esocial/base/remover", methods=["POST"])
def esocial_remover_base():
    _clear_esocial_base_session()
    return jsonify({"ok": True})


@app.route("/esocial/processar", methods=["POST"])
def esocial_processar():
    base_file = request.files.get("base_file")
    export_files = request.files.getlist("rel_files")
    base_sheet = request.form.get("base_sheet", "").strip()
    stored_base = _get_esocial_base_session_state()

    if base_file and base_file.filename:
        ok, msg = validate_uploaded_file(base_file, ESOCIAL_ALLOWED_EXTENSIONS, "a planilha base")
        if not ok:
            flash(msg)
            return redirect(url_for("esocial"))
    elif not stored_base:
        flash("Selecione a planilha base.")
        return redirect(url_for("esocial"))
    if not base_sheet:
        flash("Selecione a guia/mês da planilha base.")
        return redirect(url_for("esocial"))

    valid_exports = [f for f in export_files if f and f.filename and is_allowed_file(f.filename)]
    if not valid_exports:
        flash("Selecione uma ou mais planilhas de envios do eSocial (.xls ou .xlsx).")
        return redirect(url_for("esocial"))

    temp_root = Path(tempfile.mkdtemp(prefix="recibo_esocial_web_"))
    upload_dir = temp_root / "uploads"
    output_root = temp_root / "saida"
    upload_dir.mkdir(parents=True, exist_ok=True)
    output_root.mkdir(parents=True, exist_ok=True)

    try:
        if base_file and base_file.filename:
            base_path = upload_dir / secure_filename(base_file.filename)
            base_file.save(base_path)
        else:
            base_path = upload_dir / secure_filename(stored_base["filename"])
            shutil.copy2(stored_base["path"], base_path)
        export_paths = []
        for index, uploaded in enumerate(valid_exports, start=1):
            filename = secure_filename(Path(uploaded.filename).name)
            path = upload_dir / f"{index:03d}_{filename}"
            uploaded.save(path)
            export_paths.append(str(path))

        receipts_folder = Path(create_output_folder(str(output_root)))
        result = process_esocial_receipts(
            str(base_path), base_sheet, export_paths, str(receipts_folder)
        )
        zip_path = Path(create_esocial_zip(str(receipts_folder), result["month"], result["year"]))
        return send_file(zip_path, as_attachment=True, download_name=zip_path.name, mimetype="application/zip")
    except Exception as exc:
        logger.exception("Erro no processamento do Recibo eSocial")
        flash(str(exc) if isinstance(exc, (ValueError, KeyError, RuntimeError)) else "Erro ao processar os arquivos do Recibo eSocial.")
        return redirect(url_for("esocial"))


# =========================
# FILA ASSÍNCRONA PARA PROCESSAMENTOS PESADOS
# =========================
from edge_app.workers.jobs import JobManager

JOBS_DIR = os.path.join(DATA_DIR, "jobs")
job_manager = JobManager(
    JOBS_DIR,
    max_workers=int(os.environ.get("JOB_WORKERS", "2")),
    ttl_hours=int(os.environ.get("JOB_TTL_HOURS", "12")),
)


def _save_uploads_for_job(files, destination: Path, allowed_extensions: set[str], label: str):
    destination.mkdir(parents=True, exist_ok=True)
    saved = []
    for index, file in enumerate(files, start=1):
        ok, msg = validate_uploaded_file(file, allowed_extensions, label)
        if not ok:
            raise ValueError(msg)
        filename = secure_filename(Path(file.filename).name)
        path = destination / f"{index:03d}_{filename}"
        file.save(path)
        saved.append(path)
    return saved


def _memory_file_from_path(path: Path):
    data = path.read_bytes()
    mem = UploadedMemoryFile(path.name, data)
    mem.seek(0)
    return mem


@app.route("/jobs/<job_id>")
def job_page(job_id):
    job = job_manager.get(job_id)
    if not job:
        abort(404)
    return render_template("job_status.html", job_id=job_id, title=job.title)


@app.route("/jobs/<job_id>/status")
def job_status(job_id):
    job = job_manager.get(job_id)
    if not job:
        return jsonify({"status": "not_found", "progress": 100, "message": "Processamento não encontrado."}), 404
    return jsonify({
        "id": job.id,
        "title": job.title,
        "status": job.status,
        "progress": job.progress,
        "message": job.message,
        "has_download": bool(job.result_path and os.path.exists(job.result_path)),
    })


@app.route("/jobs/<job_id>/download")
def job_download(job_id):
    job = job_manager.get(job_id)
    if not job or job.status != "finished" or not job.result_path or not os.path.exists(job.result_path):
        abort(404)
    return send_file(job.result_path, as_attachment=True, download_name=job.download_name or Path(job.result_path).name)


@app.route("/relatorios/async", methods=["POST"])
def relatorios_async():
    files = [f for f in request.files.getlist("files") if f and f.filename]
    if not files:
        flash("Selecione ao menos uma planilha.")
        return redirect(url_for("relatorios"))
    try:
        mes = int(request.form.get("mes"))
        if mes < 1 or mes > 12:
            raise ValueError
    except Exception:
        flash("Mês inválido. Informe um valor entre 1 e 12.")
        return redirect(url_for("relatorios"))

    job_root = Path(tempfile.mkdtemp(prefix="job_relatorios_", dir=JOBS_DIR))
    try:
        saved_paths = salvar_planilhas_relatorios_uploads(files, job_root / "uploads")
    except ValueError as exc:
        flash(str(exc))
        return redirect(url_for("relatorios"))

    if not saved_paths:
        flash("Selecione planilhas em .xls/.xlsx ou um ZIP contendo as planilhas.")
        return redirect(url_for("relatorios"))

    def task(progress):
        progress(15, "Lendo planilhas...")
        mem_files = [_memory_file_from_path(p) for p in saved_paths]
        cadastros = carregar_empresas_cnpj_relatorios()
        wb_rel = criar_relatorio(mem_files, mes, cadastros)
        progress(45, "Gerando relatório...")
        mem_files = [_memory_file_from_path(p) for p in saved_paths]
        wb_base = criar_base(mem_files, mes, cadastros)
        out_dir = job_root / "saida"
        out_dir.mkdir(exist_ok=True)
        caminho_rel = out_dir / f"Relatorio_{nome_mes(mes)}.xlsx"
        caminho_base = out_dir / f"Base_do_Mes_{nome_mes(mes)}.xlsx"
        wb_rel.save(caminho_rel)
        wb_base.save(caminho_base)
        progress(75, "Compactando arquivos...")
        zip_path = out_dir / f"Arquivos_{nome_mes(mes)}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(caminho_rel, caminho_rel.name)
            z.write(caminho_base, caminho_base.name)
        return str(zip_path), zip_path.name

    job = job_manager.create(f"Relatórios — {nome_mes(mes)}", task)
    return redirect(url_for("job_page", job_id=job.id))



@app.route("/relatorios/comparar/preparar", methods=["POST"])
def relatorios_comparar_preparar():
    relatorio_file = request.files.get("relatorio_file")
    controle_file = request.files.get("controle_file")
    token = secrets.token_urlsafe(18).replace("-", "_")
    job_root = Path(JOBS_DIR) / f"comparacao_relatorios_{token}"
    try:
        caminho_relatorio_original = _salvar_upload_comparacao(relatorio_file, job_root, {".xlsx"}, "o Relatório gerado pelo sistema")
        caminho_controle_original = _salvar_upload_comparacao(controle_file, job_root, {".xlsx"}, "a planilha de controle")
        caminho_relatorio = job_root / "relatorio_upload.xlsx"
        caminho_controle = job_root / "controle_upload.xlsx"
        if caminho_relatorio_original != caminho_relatorio:
            caminho_relatorio_original.replace(caminho_relatorio)
        if caminho_controle_original != caminho_controle:
            caminho_controle_original.replace(caminho_controle)
        guias = listar_guias_planilha_controle(caminho_controle)
        if not guias:
            raise ValueError("A planilha de controle não possui guias disponíveis.")
        comparacao = {
            "token": token,
            "relatorio_nome": Path(caminho_relatorio).name,
            "controle_nome": Path(caminho_controle).name,
            "guias": guias,
        }
        flash("Arquivos carregados. Agora selecione as guias que deseja usar na comparação.")
        return render_template("relatorios.html", comparacao=comparacao, empresas_cnpj=listar_empresas_cnpj_relatorios())
    except Exception as exc:
        logger.exception("Erro ao preparar comparação de relatórios")
        shutil.rmtree(job_root, ignore_errors=True)
        flash(str(exc) or "Não foi possível ler os arquivos para comparação.")
        return redirect(url_for("relatorios"))


@app.route("/relatorios/comparar/executar", methods=["POST"])
def relatorios_comparar_executar():
    try:
        token = _safe_comparacao_token(request.form.get("token"))
        guias = request.form.getlist("guias")
        if not guias:
            raise ValueError("Selecione ao menos uma guia para comparar.")
        job_root = Path(JOBS_DIR) / f"comparacao_relatorios_{token}"
        if not job_root.exists():
            raise ValueError("Sessão de comparação expirada. Envie os arquivos novamente.")
        caminho_relatorio = job_root / "relatorio_upload.xlsx"
        caminho_controle = job_root / "controle_upload.xlsx"
        if not caminho_relatorio.exists() or not caminho_controle.exists():
            raise ValueError("Arquivos da comparação não foram encontrados. Envie novamente.")
        saida = job_root / "Relatorio_comparado.xlsx"
        comparar_relatorio_com_controle(caminho_relatorio, caminho_controle, guias, saida)
        return send_file(saida, as_attachment=True, download_name="Relatorio_comparado.xlsx")
    except Exception as exc:
        logger.exception("Erro ao executar comparação de relatórios")
        flash(str(exc) or "Não foi possível comparar as planilhas.")
        return redirect(url_for("relatorios"))


@app.route("/encaminhamentos/async", methods=["POST"])
def encaminhamentos_async():
    file = request.files.get("file")
    ok, msg = validate_uploaded_file(file, EXCEL_EXTENSIONS, "a planilha Base do Mês")
    if not ok:
        flash(msg)
        return redirect(url_for("encaminhamentos"))
    formato_saida = request.form.get("formato_saida", "docx")
    job_root = Path(tempfile.mkdtemp(prefix="job_encaminhamentos_", dir=JOBS_DIR))
    upload = job_root / secure_filename(file.filename)
    file.save(upload)

    def task(progress):
        progress(20, "Lendo base do mês...")
        with open(upload, "rb") as fh:
            fh.filename = upload.name
            progress(45, f"Gerando encaminhamentos em {formato_saida.upper()}...")
            zip_path = gerar_encaminhamentos(fh, formato_saida=formato_saida)
        final_path = job_root / "encaminhamentos.zip"
        shutil.copy2(zip_path, final_path)
        progress(90, "Finalizando pacote...")
        return str(final_path), "encaminhamentos.zip"

    job = job_manager.create("Encaminhamentos", task)
    return redirect(url_for("job_page", job_id=job.id))


@app.route("/esocial/processar/async", methods=["POST"])
def esocial_processar_async():
    wants_json = (
        request.headers.get("X-Requested-With", "").lower() == "xmlhttprequest"
        or "application/json" in request.headers.get("Accept", "").lower()
    )

    def fail(message: str, status_code: int = 400):
        if wants_json:
            return jsonify({"ok": False, "error": message}), status_code
        flash(message)
        return redirect(url_for("esocial"))

    base_file = request.files.get("base_file")
    export_files = [f for f in request.files.getlist("rel_files") if f and f.filename]
    base_sheet = request.form.get("base_sheet", "").strip()
    stored_base = _get_esocial_base_session_state()

    if base_file and base_file.filename:
        ok, msg = validate_uploaded_file(base_file, ESOCIAL_ALLOWED_EXTENSIONS, "a planilha base")
        if not ok:
            return fail(msg)
    elif not stored_base:
        return fail("Selecione a planilha base.")
    if not base_sheet:
        return fail("Selecione a guia/mês da planilha base.")
    if stored_base and not (base_file and base_file.filename) and base_sheet not in stored_base.get("sheets", []):
        return fail("A guia selecionada não pertence à planilha base atual. Selecione novamente o mês.")

    valid_exports = [f for f in export_files if is_allowed_file(f.filename)]
    if not valid_exports:
        return fail("Selecione uma ou mais planilhas de envios do eSocial (.xls ou .xlsx).")

    job_root = Path(tempfile.mkdtemp(prefix="job_recibo_esocial_", dir=JOBS_DIR))
    upload_dir = job_root / "uploads"
    output_root = job_root / "saida"
    upload_dir.mkdir(parents=True, exist_ok=True)
    output_root.mkdir(parents=True, exist_ok=True)

    try:
        if base_file and base_file.filename:
            base_path = upload_dir / secure_filename(base_file.filename)
            base_file.save(base_path)
        else:
            base_path = upload_dir / secure_filename(stored_base["filename"])
            shutil.copy2(stored_base["path"], base_path)
        export_paths = _save_uploads_for_job(
            valid_exports,
            upload_dir / "envios_esocial",
            ESOCIAL_ALLOWED_EXTENSIONS,
            "as planilhas de envios do eSocial",
        )
    except Exception as exc:
        logger.exception("Erro ao salvar uploads do Recibo eSocial")
        shutil.rmtree(job_root, ignore_errors=True)
        return fail(str(exc) or "Não foi possível preparar os arquivos enviados.")

    def task(progress):
        progress(8, "Lendo a guia selecionada e localizando OK E-SOCIAL...")
        receipts_folder = Path(create_output_folder(str(output_root)))
        result = process_esocial_receipts(
            str(base_path),
            base_sheet,
            [str(path) for path in export_paths],
            str(receipts_folder),
            progress=progress,
        )
        if result.get("total_generated", 0) == 0:
            progress(90, "Nenhum funcionário coincidiu entre a base e as planilhas enviadas. Preparando relatório de conferência...")
        else:
            progress(90, f"{result['total_generated']} recibo(s) gerado(s). Preparando o ZIP...")
        zip_path = Path(create_esocial_zip(str(receipts_folder), result["month"], result["year"]))

        total_generated = int(result.get("total_generated", len(result.get("generated", []))))
        total_without_match = int(result.get("total_without_match", 0))
        if total_generated == 0:
            details = []
            for item in result.get("summary", [])[:3]:
                details.append(
                    f"{item.get('empresa', 'Empresa')}: {item.get('total_base', 0)} na base, "
                    f"{item.get('total_export', 0)} na exportação e {item.get('total_encontrado', 0)} correspondência(s)"
                )
            detail_text = "; ".join(details)
            final_message = (
                "Processamento concluído sem PDF de recibo. "
                + (detail_text + ". " if detail_text else "")
                + "O ZIP contém o resumo detalhado com os nomes não encontrados."
            )
        elif total_without_match:
            final_message = (
                f"Concluído: {total_generated} PDF(s) gerado(s). "
                f"{total_without_match} empresa(s) ficaram sem correspondência; consulte o resumo no ZIP."
            )
        else:
            final_message = f"Concluído: {total_generated} PDF(s) de recibo gerado(s)."
        return str(zip_path), zip_path.name, final_message

    month, year = _sheet_month_year(base_sheet)
    job_label = f"Recibo eSocial - {month or base_sheet}" + (f" {year}" if year else "")
    job = job_manager.create(job_label, task)

    if wants_json:
        return jsonify({
            "ok": True,
            "job_id": job.id,
            "title": job.title,
            "status_url": url_for("job_status", job_id=job.id),
            "download_url": url_for("job_download", job_id=job.id),
        }), 202
    return redirect(url_for("job_page", job_id=job.id))

@app.errorhandler(403)
def erro_permissao(_error):
    logger.warning("Acesso negado endpoint=%s request_id=%s", request.endpoint, getattr(request, "request_id", ""))
    return render_template("erro.html", codigo=403, titulo="Acesso bloqueado", mensagem="A sessão expirou ou você não tem permissão para executar esta ação. Volte e tente novamente."), 403

@app.errorhandler(404)
def erro_nao_encontrado(_error):
    return render_template("erro.html", codigo=404, titulo="Página não encontrada", mensagem="O endereço acessado não existe ou foi movido."), 404

@app.errorhandler(413)
def arquivo_muito_grande(_error):
    return render_template("erro.html", codigo=413, titulo="Arquivo muito grande", mensagem=f"O arquivo excede o limite configurado de {get_max_upload_mb()} MB. Divida o processamento em lotes menores."), 413

@app.errorhandler(500)
def erro_interno(_error):
    logger.exception("Erro interno não tratado request_id=%s", getattr(request, "request_id", ""))
    return render_template("erro.html", codigo=500, titulo="Erro interno", mensagem="Ocorreu uma falha inesperada. Tente novamente. Se continuar, consulte os logs do Render."), 500

if __name__ == "__main__":
    init_fisico_db()
    app.run(debug=False, host="0.0.0.0", port=5000)
