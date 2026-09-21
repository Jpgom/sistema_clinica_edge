# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import json
import os
import re
import shutil
import threading
import secrets
import time
import uuid
import webbrowser
import zipfile
from dataclasses import asdict
from datetime import timedelta
from pathlib import Path
from typing import Any

from flask import Flask, abort, jsonify, redirect, render_template, request, send_file, session, url_for
from werkzeug.utils import secure_filename

from . import core
from .archive_store import ArchiveStore, digits_only, format_document

APP_DIR = Path(__file__).resolve().parent

# Dados persistentes da função Separar exames.
# No Render, o app.py principal define SEPARAR_EXAMES_DATA_DIR dentro do disco /var/data.
def _default_data_dir() -> Path:
    configured = (os.environ.get("SEPARAR_EXAMES_DATA_DIR") or os.environ.get("EDGE_EXTRATOR_DATA_DIR") or "").strip()
    if configured:
        return Path(configured)
    persist_root = (os.environ.get("RENDER_DISK_PATH") or os.environ.get("DATA_DIR") or "").strip()
    if persist_root:
        return Path(persist_root) / "separar_exames"
    return APP_DIR / "dados"

DATA_DIR = _default_data_dir()
BUNDLED_DATA_DIR = APP_DIR / "dados"
WORK_DIR = Path(os.environ.get("SEPARAR_EXAMES_WORK_DIR") or (DATA_DIR / "trabalho_web"))
LIST_DIR = WORK_DIR / "listas"
JOBS_DIR = WORK_DIR / "jobs"
CONFIG_FILE = DATA_DIR / "web_config.json"
UNITS_FILE = DATA_DIR / "unidades.json"
SECRET_FILE = WORK_DIR / ".web_secret"
JOB_STATE_FILENAME = "job_state.json"

def _prepare_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    (DATA_DIR / "modelos").mkdir(parents=True, exist_ok=True)
    (DATA_DIR / "tessdata").mkdir(parents=True, exist_ok=True)
    bundled_models = BUNDLED_DATA_DIR / "modelos.json"
    target_models = DATA_DIR / "modelos.json"
    if bundled_models.exists() and not target_models.exists():
        try:
            shutil.copy2(bundled_models, target_models)
        except Exception:
            target_models.write_text("[]", encoding="utf-8")

_prepare_data_dir()

# Faz o core usar os mesmos caminhos persistentes.
core.DATA_DIR = DATA_DIR
core.MODELS_DIR = DATA_DIR / "modelos"
core.MODELS_JSON = DATA_DIR / "modelos.json"
core.TESSDATA_DIR = DATA_DIR / "tessdata"
core.CONFIG_JSON = DATA_DIR / "config.json"
core.MODELS_DIR.mkdir(parents=True, exist_ok=True)
if not core.MODELS_JSON.exists():
    core.MODELS_JSON.write_text("[]", encoding="utf-8")
core.refresh_exam_types()

def _default_archive_dir() -> Path:
    configured = (os.environ.get("SEPARAR_EXAMES_ARCHIVE_DIR") or os.environ.get("EDGE_ARCHIVE_DIR") or "").strip()
    if configured:
        return Path(configured)
    return DATA_DIR / "arquivo_exames"

ARCHIVE_DIR = _default_archive_dir()
ARCHIVE = ArchiveStore(ARCHIVE_DIR)

for p in (WORK_DIR, LIST_DIR, JOBS_DIR):
    p.mkdir(parents=True, exist_ok=True)

def _persistent_secret_key() -> str:
    env_key = os.environ.get("EDGE_SECRET_KEY", "").strip()
    if env_key:
        return env_key
    try:
        if SECRET_FILE.exists():
            value = SECRET_FILE.read_text(encoding="utf-8").strip()
            if value:
                return value
        SECRET_FILE.parent.mkdir(parents=True, exist_ok=True)
        value = secrets.token_hex(32)
        SECRET_FILE.write_text(value, encoding="utf-8")
        return value
    except Exception:
        return secrets.token_hex(32)

app = Flask(__name__)
app.secret_key = _persistent_secret_key()
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024 * 1024  # 1 GB por requisição
app.config["JSON_AS_ASCII"] = False
app.permanent_session_lifetime = timedelta(days=3650)

DEFAULT_CONFIG = {
    "use_ocr": True,
    # Modo confiabilidade: faz leitura completa para não deixar exames para trás.
    "fast_mode": False,
    "stop_when_complete": False,
    "auto_threshold": 68,
    "employee_threshold": 78,
}



def _unit_id(value: str) -> str:
    raw = core.normalize_for_match(value)
    raw = re.sub(r"[^A-Z0-9]+", "_", raw).strip("_")
    return raw[:50] or "GERAL"


def _load_units() -> list[dict[str, str]]:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not UNITS_FILE.exists():
        defaults = [
            {"id": "BELEM", "name": "Belém"},
            {"id": "MACAPA", "name": "Macapá"},
        ]
        UNITS_FILE.write_text(json.dumps(defaults, ensure_ascii=False, indent=2), encoding="utf-8")
        return defaults
    try:
        raw = json.loads(UNITS_FILE.read_text(encoding="utf-8"))
        items = raw.get("units") if isinstance(raw, dict) else raw
        units: list[dict[str, str]] = []
        seen: set[str] = set()
        for item in items or []:
            name = str(item.get("name") if isinstance(item, dict) else item or "").strip()
            uid = str(item.get("id") if isinstance(item, dict) else _unit_id(name)).strip() or _unit_id(name)
            if name and uid and uid not in seen:
                units.append({"id": uid, "name": name})
                seen.add(uid)
        if units:
            return units
    except Exception:
        pass
    return [{"id": "GERAL", "name": "Geral"}]


def _save_units(units: list[dict[str, str]]) -> None:
    clean: list[dict[str, str]] = []
    seen: set[str] = set()
    for item in units:
        name = str(item.get("name") or "").strip()
        uid = str(item.get("id") or _unit_id(name)).strip() or _unit_id(name)
        if name and uid and uid not in seen:
            clean.append({"id": uid, "name": name})
            seen.add(uid)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    UNITS_FILE.write_text(json.dumps(clean, ensure_ascii=False, indent=2), encoding="utf-8")


def _unit_by_id(unit_id: str) -> dict[str, str] | None:
    unit_id = str(unit_id or "").strip()
    for u in _load_units():
        if u["id"] == unit_id:
            return u
    return None

JOBS: dict[str, dict[str, Any]] = {}
JOBS_LOCK = threading.RLock()


def _employee_to_dict(employee: core.ExpectedEmployee) -> dict[str, Any]:
    return {
        "row_id": employee.row_id,
        "name": employee.name,
        "cpf": employee.cpf,
        "company": employee.company,
        "cnpj": employee.cnpj,
        "expected_exams": sorted(employee.expected_exams),
        "expected_exam_counts": dict(employee.expected_exam_counts),
        "expected_exam_receipts": {str(k): list(v) for k, v in employee.expected_exam_receipts.items()},
        "source_sheet": employee.source_sheet,
    }


def _employee_from_dict(data: dict[str, Any]) -> core.ExpectedEmployee:
    return core.ExpectedEmployee(
        row_id=int(data.get("row_id", 0)),
        name=str(data.get("name", "")),
        cpf=str(data.get("cpf", "")),
        company=str(data.get("company", "")),
        cnpj=str(data.get("cnpj", "")),
        expected_exams=set(data.get("expected_exams") or []),
        expected_exam_counts={str(k): int(v) for k, v in (data.get("expected_exam_counts") or {}).items()},
        expected_exam_receipts={str(k): [str(x) for x in (v or [])] for k, v in (data.get("expected_exam_receipts") or {}).items()},
        source_sheet=str(data.get("source_sheet", "")),
    )


def _summary_to_dict(summary: core.ProcessingSummary | None) -> dict[str, Any] | None:
    if summary is None:
        return None
    return asdict(summary)


def _summary_from_dict(data: dict[str, Any] | None, workspace: Path, output_dir: Path) -> core.ProcessingSummary | None:
    if not data:
        return None
    analyses: list[core.PageAnalysis] = []
    upload_dir = workspace / "uploads"
    for raw in data.get("analyses") or []:
        item = dict(raw)
        source_name = Path(str(item.get("source_pdf_path") or item.get("source_pdf") or "")).name
        candidate = upload_dir / source_name
        if candidate.exists():
            item["source_pdf_path"] = str(candidate)
        analyses.append(core.PageAnalysis(**item))
    payload = dict(data)
    payload["root_dir"] = str(output_dir)
    payload["analyses"] = analyses
    return core.ProcessingSummary(**payload)


def _persist_job(job: dict[str, Any]) -> None:
    if job.get("summary") is None:
        return
    workspace = Path(job["workspace"])
    workspace.mkdir(parents=True, exist_ok=True)
    state = {
        "format": 1,
        "id": job["id"],
        "status": job.get("status", "done"),
        "progress": job.get("progress", 100),
        "current": job.get("current", 0),
        "total": job.get("total", 0),
        "message": job.get("message", ""),
        "error": job.get("error", ""),
        "created_at": job.get("created_at", time.time()),
        "folder_name": job.get("folder_name", "ARQUIVOS SEPARADOS"),
        "config": job.get("config") or DEFAULT_CONFIG.copy(),
        "pdf_files": [Path(p).name for p in job.get("pdf_paths", [])],
        "employees": [_employee_to_dict(e) for e in job.get("employees", [])],
        "summary": _summary_to_dict(job.get("summary")),
    }
    target = workspace / JOB_STATE_FILENAME
    temp = workspace / (JOB_STATE_FILENAME + ".tmp")
    temp.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
    os.replace(temp, target)


def _load_persisted_job(job_id: str) -> dict[str, Any] | None:
    workspace = JOBS_DIR / job_id
    state_file = workspace / JOB_STATE_FILENAME
    if not state_file.is_file():
        return None
    try:
        data = json.loads(state_file.read_text(encoding="utf-8"))
        folder_name = _safe_folder_name(str(data.get("folder_name") or "ARQUIVOS SEPARADOS"))
        output_dir = workspace / "saida" / folder_name
        upload_dir = workspace / "uploads"
        pdf_paths = [upload_dir / Path(str(name)).name for name in (data.get("pdf_files") or [])]
        employees = [_employee_from_dict(x) for x in (data.get("employees") or [])]
        summary = _summary_from_dict(data.get("summary"), workspace, output_dir)
        if summary is None:
            return None
        return {
            "id": job_id,
            "status": "done",
            "progress": 100,
            "current": int(data.get("current", 0)),
            "total": int(data.get("total", 0)),
            "message": str(data.get("message") or "Processamento concluído."),
            "error": str(data.get("error") or ""),
            "created_at": float(data.get("created_at", time.time())),
            "folder_name": folder_name,
            "workspace": workspace,
            "output_dir": output_dir,
            "pdf_paths": pdf_paths,
            "employees": employees,
            "config": data.get("config") or DEFAULT_CONFIG.copy(),
            "summary": summary,
        }
    except Exception:
        return None


def _restore_persisted_jobs() -> None:
    with JOBS_LOCK:
        for folder in JOBS_DIR.iterdir():
            if not folder.is_dir():
                continue
            job = _load_persisted_job(folder.name)
            if job:
                JOBS[folder.name] = job


def _load_config() -> dict[str, Any]:
    cfg = DEFAULT_CONFIG.copy()
    try:
        if CONFIG_FILE.exists():
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                cfg.update({k: data[k] for k in cfg if k in data})
    except Exception:
        pass
    cfg["use_ocr"] = bool(cfg.get("use_ocr", True))
    cfg["fast_mode"] = False
    # Modo confiável: sempre lê o lote inteiro; mantém a chave apenas por compatibilidade.
    cfg["stop_when_complete"] = False
    cfg["auto_threshold"] = max(50, min(85, int(float(cfg.get("auto_threshold", 68)))))
    cfg["employee_threshold"] = max(50, min(90, int(float(cfg.get("employee_threshold", 78)))))
    return cfg


def _save_config(cfg: dict[str, Any]) -> None:
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def _safe_folder_name(value: str) -> str:
    value = (value or "ARQUIVOS SEPARADOS").strip()
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", value)
    value = re.sub(r"\s+", " ", value).strip(" .")
    return value[:90] or "ARQUIVOS SEPARADOS"


def _unique_upload_path(folder: Path, name: str) -> Path:
    base = secure_filename(name) or "arquivo.pdf"
    target = folder / base
    if not target.exists():
        return target
    stem, suffix = target.stem, target.suffix
    n = 2
    while True:
        candidate = folder / f"{stem}_{n}{suffix}"
        if not candidate.exists():
            return candidate
        n += 1


def _serialize_analysis(a: core.PageAnalysis) -> dict[str, Any]:
    return {
        "id": a.id,
        "source_pdf": a.source_pdf,
        "page_number": a.page_number,
        "exam_type": a.exam_type,
        "exam_subtype": a.exam_subtype,
        "employee_name": a.employee_name,
        "company": a.company,
        "cnpj": a.cnpj,
        "exam_confidence": a.exam_confidence,
        "employee_match_confidence": a.employee_match_confidence,
        "status": a.status,
        "reason": a.reason,
        "output_file": a.output_file,
        "expected": a.expected,
        "employee_row_id": a.employee_row_id,
    }


def _summary_payload(summary: core.ProcessingSummary) -> dict[str, Any]:
    relevant = [
        _serialize_analysis(a)
        for a in summary.analyses
        if a.status not in {"IGNORADO", "IGNORADO_MANUAL"}
    ]
    saved = sum(a.status in {"SALVO_AUTOMATICO", "SALVO_MANUAL", "ANEXADO_CONTINUACAO"} for a in summary.analyses)
    return {
        "root_dir": summary.root_dir,
        "total_pages": summary.total_pages,
        "saved": saved,
        "pending": summary.pending,
        "duplicates": summary.duplicates,
        "missing": summary.missing_expected,
        "ignored": summary.ignored,
        "unexpected": summary.unexpected,
        "quick_scanned": summary.quick_scanned,
        "detailed_ocr_pages": summary.detailed_ocr_pages,
        "skipped_after_complete": summary.skipped_after_complete,
        "fast_rejected": getattr(summary, "fast_rejected", 0),
        "rescued_pages": getattr(summary, "rescued_pages", 0),
        "audit_gaps": getattr(summary, "audit_gaps", 0),
        "analyses": relevant,
        "missing_rows": summary.missing_rows,
    }


def _archive_candidates(job: dict[str, Any]) -> list[dict[str, Any]]:
    summary: core.ProcessingSummary | None = job.get("summary")
    if summary is None:
        return []
    root = Path(job["output_dir"])
    saved_status = {"SALVO_AUTOMATICO", "SALVO_MANUAL", "ANEXADO_CONTINUACAO"}
    by_output: dict[str, core.PageAnalysis] = {}
    for a in summary.analyses:
        if a.status not in saved_status or not a.output_file:
            continue
        current = by_output.get(a.output_file)
        # Prefere a análise principal, que normalmente contém os metadados mais completos.
        if current is None or (not current.employee_name and a.employee_name) or current.status == "ANEXADO_CONTINUACAO":
            by_output[a.output_file] = a

    employees = {int(e.row_id): e for e in job.get("employees", [])}
    docs: list[dict[str, Any]] = []
    for rel, a in by_output.items():
        src = (root / rel).resolve()
        if not src.is_file() or (root.resolve() not in src.parents and src != root.resolve()):
            continue
        employee = employees.get(int(a.employee_row_id)) if a.employee_row_id is not None else None
        company_name = (a.company or (employee.company if employee else "") or "").strip()
        company_doc = digits_only(a.cnpj or (employee.cnpj if employee else "") or "")
        employee_name = (a.employee_name or (employee.name if employee else "") or "").strip()
        employee_cpf = digits_only(a.employee_cpf or (employee.cpf if employee else "") or "")
        docs.append({
            "source_path": str(src),
            "company_name": company_name,
            "company_document": company_doc,
            "employee_name": employee_name,
            "employee_cpf": employee_cpf,
            "exam_type": (a.exam_type or "OUTROS").upper(),
            "exam_subtype": (a.exam_subtype or "").upper(),
            "receipt": (getattr(a, "receipt", "") or "").strip(),
            "source_analysis_id": a.id,
        })
    return docs


def _job_public(job: dict[str, Any]) -> dict[str, Any]:
    data = {
        "id": job["id"],
        "status": job.get("status", "queued"),
        "progress": job.get("progress", 0),
        "current": job.get("current", 0),
        "total": job.get("total", 0),
        "message": job.get("message", ""),
        "error": job.get("error", ""),
        "folder_name": job.get("folder_name", "ARQUIVOS SEPARADOS"),
    }
    summary = job.get("summary")
    if summary is not None:
        data["summary"] = _summary_payload(summary)
        data["download_zip"] = url_for("download_zip", job_id=job["id"])
        data["download_report"] = url_for("download_report", job_id=job["id"])
        data["review_url"] = url_for("review_page", job_id=job["id"])
        data["archive_eligible"] = len(_archive_candidates(job))
        data["archive_saved"] = ARCHIVE.count_for_job(job["id"])
    return data


def _get_job(job_id: str) -> dict[str, Any]:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            job = _load_persisted_job(job_id)
            if job:
                JOBS[job_id] = job
        if not job:
            abort(404)
        return job


def _cleanup_old_workspaces(max_age_hours: int = 24) -> None:
    # Somente uploads temporários de planilhas são limpos automaticamente.
    # Resultados de extrações ficam disponíveis até o usuário clicar em LIMPAR.
    cutoff = time.time() - max_age_hours * 3600
    for p in LIST_DIR.iterdir():
        try:
            if p.stat().st_mtime < cutoff:
                if p.is_dir():
                    shutil.rmtree(p, ignore_errors=True)
                else:
                    p.unlink(missing_ok=True)
        except Exception:
            pass


def _run_job(job_id: str) -> None:
    job = _get_job(job_id)
    try:
        with JOBS_LOCK:
            job["status"] = "processing"
            job["message"] = "Preparando os arquivos..."

        cfg = job["config"]

        def progress(cur: int, total: int, msg: str) -> None:
            with JOBS_LOCK:
                job["current"] = cur
                job["total"] = total
                job["progress"] = round((cur / total * 100), 1) if total else 0
                job["message"] = msg

        summary = core.process_pdfs(
            job["pdf_paths"],
            job["employees"],
            job["output_dir"],
            use_ocr=cfg["use_ocr"],
            auto_threshold=float(cfg["auto_threshold"]),
            employee_threshold=float(cfg["employee_threshold"]),
            progress=progress,
            fast_mode=cfg["fast_mode"],
            stop_when_complete=cfg["stop_when_complete"],
        )
        with JOBS_LOCK:
            job["summary"] = summary
            job["status"] = "done"
            job["progress"] = 100
            job["message"] = "Processamento concluído."
            _persist_job(job)
    except Exception as exc:
        with JOBS_LOCK:
            job["status"] = "error"
            job["error"] = str(exc)
            job["message"] = "Falha no processamento."


def _make_zip(job: dict[str, Any]) -> Path:
    output_dir: Path = Path(job["output_dir"])
    zip_path = Path(job["workspace"]) / f"{job['folder_name']}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in output_dir.rglob("*"):
            if not p.is_file() or p.name == ".edge_extrator_output":
                continue
            arcname = Path(job["folder_name"]) / p.relative_to(output_dir)
            zf.write(p, arcname.as_posix())
    return zip_path


# Restaura extrações concluídas gravadas em disco. Elas só são removidas pelo botão LIMPAR.
_restore_persisted_jobs()


@app.before_request
def optional_login_guard():
    session.permanent = True
    password = os.environ.get("EDGE_WEB_PASSWORD", "").strip()
    if not password:
        return None
    if request.endpoint in {"login", "health", "static"}:
        return None
    if session.get("edge_auth") is True:
        return None
    if request.path.startswith("/api/") or request.path.startswith("/download/"):
        return jsonify({"error": "Sessão expirada. Entre novamente no sistema."}), 401
    return redirect(url_for("login", next=request.path))


@app.route("/login", methods=["GET", "POST"])
def login():
    password = os.environ.get("EDGE_WEB_PASSWORD", "").strip()
    if not password:
        return redirect(url_for("index"))
    error = ""
    if request.method == "POST":
        if request.form.get("password", "") == password:
            session["edge_auth"] = True
            target = request.args.get("next") or url_for("index")
            if not str(target).startswith("/"):
                target = url_for("index")
            return redirect(target)
        error = "Senha incorreta."
    return render_template("login.html", error=error, version="5.0 WEB")


@app.post("/logout")
def logout():
    active_job_id = session.get("active_job_id")
    session.clear()
    if active_job_id:
        session["active_job_id"] = active_job_id
    return redirect(url_for("login"))


@app.get("/")
def index():
    requested_job = (request.args.get("job") or "").strip()
    if requested_job:
        try:
            _get_job(requested_job)
            session["active_job_id"] = requested_job
        except Exception:
            pass
    cfg = _load_config()
    return render_template(
        "index.html",
        version="5.0 WEB",
        config=cfg,
        ocr_available=bool(core.locate_tesseract()),
    )


@app.get("/configuracoes")
def settings_page():
    return render_template(
        "settings.html",
        version="5.0 WEB",
        config=_load_config(),
        tesseract=core.locate_tesseract() or "Não localizado",
    )


@app.get("/modelos")
def models_page():
    return render_template("models.html", version="5.0 WEB", exam_types=core.get_exam_types())


@app.get("/arquivo")
def archive_page():
    return render_template("archive.html", version="5.0 WEB")


@app.get("/revisao/<job_id>")
def review_page(job_id: str):
    _get_job(job_id)
    session["active_job_id"] = job_id
    return render_template("review.html", version="5.0 WEB", job_id=job_id, exam_types=core.get_exam_types())


@app.get("/api/config")
def api_get_config():
    cfg = _load_config()
    return jsonify({**cfg, "ocr_available": bool(core.locate_tesseract()), "tesseract": core.locate_tesseract() or ""})


@app.post("/api/config")
def api_save_config():
    data = request.get_json(silent=True) or request.form
    cfg = {
        "use_ocr": str(data.get("use_ocr", "true")).lower() in {"1", "true", "on", "yes"},
        "fast_mode": False,
        "stop_when_complete": False,
        "auto_threshold": max(50, min(85, int(float(data.get("auto_threshold", 68))))),
        "employee_threshold": max(50, min(90, int(float(data.get("employee_threshold", 78))))),
    }
    _save_config(cfg)
    return jsonify({"ok": True, "config": cfg})


@app.post("/api/planilha")
def api_upload_sheet():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "Selecione uma planilha."}), 400
    if Path(f.filename).suffix.lower() not in {".xlsx", ".xlsm", ".csv"}:
        return jsonify({"error": "Formato de planilha não suportado."}), 400
    token = uuid.uuid4().hex
    folder = LIST_DIR / token
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / (secure_filename(f.filename) or "lista.xlsx")
    f.save(path)
    try:
        sheets = core.spreadsheet_sheets(path)
        sheet = sheets[0] if sheets else ""
        employees = core.load_expected_list(path, sheet or None)
        return jsonify({
            "token": token,
            "filename": f.filename,
            "sheets": sheets,
            "selected_sheet": sheet,
            "employees": len(employees),
            "exams": sum(e.expected_total for e in employees),
        })
    except Exception as exc:
        shutil.rmtree(folder, ignore_errors=True)
        return jsonify({"error": str(exc)}), 400


@app.get("/api/planilha/<token>/resumo")
def api_sheet_summary(token: str):
    folder = LIST_DIR / token
    if not folder.exists():
        return jsonify({"error": "Planilha temporária expirada. Envie novamente."}), 404
    files = [p for p in folder.iterdir() if p.is_file()]
    if not files:
        return jsonify({"error": "Planilha não encontrada."}), 404
    sheet = request.args.get("sheet") or None
    try:
        employees = core.load_expected_list(files[0], sheet)
        return jsonify({"employees": len(employees), "exams": sum(e.expected_total for e in employees)})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/jobs")
def api_create_job():
    active_id = str(session.get("active_job_id") or "").strip()
    if active_id:
        try:
            active_job = _get_job(active_id)
        except Exception:
            session.pop("active_job_id", None)
        else:
            if active_job.get("status") == "error":
                with JOBS_LOCK:
                    JOBS.pop(active_id, None)
                shutil.rmtree(Path(active_job["workspace"]), ignore_errors=True)
                session.pop("active_job_id", None)
            else:
                return jsonify({"error": "Existe uma extração ativa. Use 'Limpar extração' antes de iniciar um novo lote."}), 409

    token = (request.form.get("list_token") or "").strip()
    sheet = (request.form.get("sheet") or "").strip() or None
    folder_name = _safe_folder_name(request.form.get("folder_name") or "ARQUIVOS SEPARADOS")
    list_folder = LIST_DIR / token
    if not token or not list_folder.exists():
        return jsonify({"error": "Carregue a lista de funcionários antes de processar."}), 400
    list_files = [p for p in list_folder.iterdir() if p.is_file()]
    if not list_files:
        return jsonify({"error": "Planilha temporária não encontrada."}), 400

    uploads = request.files.getlist("pdfs")
    uploads = [f for f in uploads if f and f.filename]
    if not uploads:
        return jsonify({"error": "Adicione pelo menos um PDF."}), 400
    for f in uploads:
        if Path(f.filename).suffix.lower() != ".pdf":
            return jsonify({"error": f"O arquivo {f.filename} não é PDF."}), 400

    try:
        employees = core.load_expected_list(list_files[0], sheet)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    if not employees:
        return jsonify({"error": "A lista não contém funcionários/exames válidos."}), 400

    job_id = uuid.uuid4().hex[:12]
    workspace = JOBS_DIR / job_id
    upload_dir = workspace / "uploads"
    output_dir = workspace / "saida" / folder_name
    upload_dir.mkdir(parents=True, exist_ok=True)
    pdf_paths: list[Path] = []
    for f in uploads:
        p = _unique_upload_path(upload_dir, f.filename)
        f.save(p)
        pdf_paths.append(p)

    job = {
        "id": job_id,
        "status": "queued",
        "progress": 0,
        "current": 0,
        "total": 0,
        "message": "Na fila...",
        "error": "",
        "created_at": time.time(),
        "folder_name": folder_name,
        "workspace": workspace,
        "output_dir": output_dir,
        "pdf_paths": pdf_paths,
        "employees": employees,
        "config": _load_config(),
        "summary": None,
    }
    with JOBS_LOCK:
        JOBS[job_id] = job
    session["active_job_id"] = job_id
    threading.Thread(target=_run_job, args=(job_id,), daemon=True).start()
    return jsonify({"job_id": job_id, "status_url": url_for("api_job", job_id=job_id)})


@app.get("/api/jobs/active")
def api_active_job():
    job_id = (session.get("active_job_id") or "").strip()
    if not job_id:
        return jsonify({"active": False})
    try:
        job = _get_job(job_id)
    except Exception:
        session.pop("active_job_id", None)
        return jsonify({"active": False})
    with JOBS_LOCK:
        return jsonify({"active": True, "job": _job_public(job)})


@app.get("/api/jobs/<job_id>")
def api_job(job_id: str):
    job = _get_job(job_id)
    session["active_job_id"] = job_id
    with JOBS_LOCK:
        return jsonify(_job_public(job))


@app.delete("/api/jobs/<job_id>")
def api_clear_job(job_id: str):
    job = _get_job(job_id)
    if job.get("status") in {"queued", "processing"}:
        return jsonify({"error": "Aguarde o processamento terminar antes de limpar."}), 409
    workspace = Path(job["workspace"])
    with JOBS_LOCK:
        JOBS.pop(job_id, None)
    if session.get("active_job_id") == job_id:
        session.pop("active_job_id", None)
    shutil.rmtree(workspace, ignore_errors=True)
    return jsonify({"ok": True})


@app.get("/api/jobs/<job_id>/review")
def api_review(job_id: str):
    job = _get_job(job_id)
    if job.get("summary") is None:
        return jsonify({"error": "O processamento ainda não terminou."}), 409
    summary: core.ProcessingSummary = job["summary"]
    items = [
        _serialize_analysis(a)
        for a in summary.analyses
        if a.status in {"PENDENTE", "DUPLICADO"}
    ]
    employees = [
        {
            "row_id": e.row_id,
            "name": e.name,
            "company": e.company,
            "expected_exams": sorted(e.expected_exams),
        }
        for e in job["employees"]
    ]
    return jsonify({"items": items, "employees": employees})


@app.post("/api/jobs/<job_id>/review/<analysis_id>")
def api_resolve_review(job_id: str, analysis_id: str):
    job = _get_job(job_id)
    summary: core.ProcessingSummary | None = job.get("summary")
    if summary is None:
        return jsonify({"error": "O processamento ainda não terminou."}), 409
    current_item = next((a for a in summary.analyses if a.id == analysis_id), None)
    if current_item is None:
        return jsonify({"error": "Pendência não encontrada."}), 404
    if current_item.status not in {"PENDENTE", "DUPLICADO"}:
        return jsonify({"ok": True, "already_resolved": True, "job": _job_public(job)})
    data = request.get_json(silent=True) or {}
    action = str(data.get("action", "SALVAR")).upper()
    if action == "IGNORAR":
        try:
            core.resolve_pending(summary, analysis_id, "ASO", job["employees"][0], action="IGNORAR", employees=job["employees"])
            with JOBS_LOCK:
                _persist_job(job)
            return jsonify({"ok": True, "job": _job_public(job)})
        except Exception as exc:
            return jsonify({"error": str(exc)}), 400

    exam_type = core.normalize_exam(str(data.get("exam_type", "")))
    try:
        row_id = int(data.get("employee_row_id"))
    except Exception:
        return jsonify({"error": "Selecione o funcionário."}), 400
    employee = core.employees_by_row(job["employees"], row_id)
    if not employee:
        return jsonify({"error": "Funcionário não encontrado."}), 400
    if not exam_type:
        return jsonify({"error": "Selecione o tipo de exame."}), 400
    try:
        core.resolve_pending(summary, analysis_id, exam_type, employee, action="SALVAR", employees=job["employees"])
        with JOBS_LOCK:
            _persist_job(job)
        return jsonify({"ok": True, "job": _job_public(job)})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/jobs/<job_id>/page/<analysis_id>.png")
def api_review_image(job_id: str, analysis_id: str):
    job = _get_job(job_id)
    summary: core.ProcessingSummary | None = job.get("summary")
    if summary is None:
        abort(404)
    item = next((a for a in summary.analyses if a.id == analysis_id), None)
    if not item:
        abort(404)
    import pymupdf as fitz
    doc = fitz.open(item.source_pdf_path)
    try:
        page = doc[item.page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(1.35, 1.35), alpha=False)
        data = pix.tobytes("png")
    finally:
        doc.close()
    return send_file(io.BytesIO(data), mimetype="image/png", max_age=0)


@app.get("/download/<job_id>/zip")
def download_zip(job_id: str):
    job = _get_job(job_id)
    if job.get("summary") is None:
        abort(409)
    zip_path = _make_zip(job)
    return send_file(zip_path, as_attachment=True, download_name=zip_path.name, max_age=0)


@app.get("/download/<job_id>/report")
def download_report(job_id: str):
    job = _get_job(job_id)
    if job.get("summary") is None:
        abort(409)
    report = Path(job["output_dir"]) / core.REPORT_FILENAME
    if not report.exists():
        abort(404)
    return send_file(report, as_attachment=True, download_name=core.REPORT_FILENAME, max_age=0)


@app.get("/download/<job_id>/file/<path:filename>")
def download_result_file(job_id: str, filename: str):
    job = _get_job(job_id)
    root = Path(job["output_dir"]).resolve()
    target = (root / filename).resolve()
    if root not in target.parents and target != root:
        abort(403)
    if not target.is_file():
        abort(404)
    return send_file(target, as_attachment=True, download_name=target.name, max_age=0)


@app.post("/api/jobs/<job_id>/archive")
def api_archive_job(job_id: str):
    job = _get_job(job_id)
    if job.get("summary") is None:
        return jsonify({"error": "O processamento ainda não terminou."}), 409
    data = request.get_json(silent=True) or request.form
    try:
        month = int(data.get("month"))
        year = int(data.get("year"))
    except Exception:
        return jsonify({"error": "Escolha o mês e o ano antes de salvar no arquivo."}), 400
    if month < 1 or month > 12 or year < 2000 or year > 2100:
        return jsonify({"error": "Competência inválida."}), 400
    unit_id = str(data.get("unit_id") or data.get("unit") or "").strip()
    unit = _unit_by_id(unit_id)
    if not unit:
        return jsonify({"error": "Escolha a unidade antes de salvar no arquivo."}), 400
    docs = _archive_candidates(job)
    if not docs:
        return jsonify({"error": "Não há arquivos extraídos e aprovados para arquivar."}), 400
    try:
        result = ARCHIVE.save_documents(docs, year, month, source_job_id=job_id, unit_id=unit["id"], unit_name=unit["name"])
        result["competency"] = f"{unit['name']} - {month:02d}/{year}"
        result["archive_saved"] = ARCHIVE.count_for_job(job_id)
        result["archive_eligible"] = len(docs)
        return jsonify({"ok": True, **result})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400



@app.get("/api/unidades")
def api_units():
    archived = {u["id"]: u for u in ARCHIVE.units()}
    units = []
    for u in _load_units():
        x = dict(u)
        x["archive_count"] = int(archived.get(u["id"], {}).get("count", 0))
        units.append(x)
    return jsonify({"units": units})


@app.post("/api/unidades")
def api_add_unit():
    data = request.get_json(silent=True) or request.form
    name = str(data.get("name") or "").strip()
    if not name:
        return jsonify({"error": "Informe o nome da unidade."}), 400
    units = _load_units()
    uid = _unit_id(name)
    if any(u["id"] == uid for u in units):
        return jsonify({"error": "Esta unidade já está cadastrada."}), 400
    units.append({"id": uid, "name": name})
    _save_units(units)
    return jsonify({"ok": True, "unit": {"id": uid, "name": name}, "units": _load_units()})


@app.delete("/api/unidades/<unit_id>")
def api_delete_unit(unit_id: str):
    units = _load_units()
    unit_id = str(unit_id or "").strip()
    if unit_id in {"BELEM", "MACAPA", "GERAL"}:
        return jsonify({"error": "Esta unidade padrão não pode ser excluída."}), 400
    if any(u.get("id") == unit_id and int(u.get("count", 0)) for u in ARCHIVE.units()):
        return jsonify({"error": "Esta unidade possui arquivos arquivados e não pode ser excluída."}), 400
    kept = [u for u in units if u["id"] != unit_id]
    if len(kept) == len(units):
        return jsonify({"error": "Unidade não encontrada."}), 404
    _save_units(kept)
    return jsonify({"ok": True, "units": _load_units()})


@app.get("/api/archive/meta")
def api_archive_meta():
    year = request.args.get("year", type=int)
    month = request.args.get("month", type=int)
    unit_id = request.args.get("unit", "")
    return jsonify({
        "units": _load_units(),
        "periods": ARCHIVE.periods(unit_id=unit_id),
        "companies": ARCHIVE.companies(year=year, month=month, unit_id=unit_id),
        "exam_types": ARCHIVE.exam_types(unit_id=unit_id),
        "all_exam_types": list(core.get_exam_types()),
        "receipt_filters": ["RECIBOS", "A PRAZO"],
    })


@app.get("/api/archive")
def api_archive_search():
    year = request.args.get("year", type=int)
    month = request.args.get("month", type=int)
    page = request.args.get("page", default=1, type=int)
    page_size = request.args.get("page_size", default=100, type=int)
    unit_id = request.args.get("unit", "")
    data = ARCHIVE.search(
        q=request.args.get("q", ""), year=year, month=month,
        company_key=request.args.get("company", ""), exam_type=request.args.get("exam_type", ""),
        receipt_filter=request.args.get("receipt_filter", ""), unit_id=unit_id,
        page=page, page_size=page_size,
    )
    return jsonify(data)


@app.delete("/api/archive/<int:doc_id>")
def api_archive_delete_one(doc_id: int):
    if not ARCHIVE.get(doc_id):
        return jsonify({"error": "Documento arquivado não encontrado."}), 404
    result = ARCHIVE.delete_documents([doc_id])
    return jsonify({"ok": True, **result})


@app.delete("/api/archive")
def api_archive_delete_many():
    data = request.get_json(silent=True) or {}
    raw_ids = data.get("ids") or []
    ids: list[int] = []
    for value in raw_ids:
        try:
            ids.append(int(value))
        except Exception:
            pass

    # Também permite excluir exatamente o resultado dos filtros visíveis.
    if not ids and data.get("use_filters"):
        rows = ARCHIVE.filtered_rows(
            q=str(data.get("q") or ""),
            year=int(data["year"]) if data.get("year") else None,
            month=int(data["month"]) if data.get("month") else None,
            company_key=str(data.get("company") or ""),
            exam_type=str(data.get("exam_type") or ""),
            receipt_filter=str(data.get("receipt_filter") or ""),
            unit_id=str(data.get("unit") or ""),
        )
        ids = [int(r.get("id", 0)) for r in rows if int(r.get("id", 0))]

    if not ids:
        return jsonify({"error": "Nenhum documento selecionado para exclusão."}), 400
    result = ARCHIVE.delete_documents(ids)
    return jsonify({"ok": True, **result})


@app.get("/arquivo/documento/<int:doc_id>/visualizar")
def archive_preview(doc_id: int):
    path = ARCHIVE.path_for(doc_id)
    if not path:
        abort(404)
    return send_file(path, mimetype="application/pdf", as_attachment=False, download_name=path.name, max_age=0)


@app.get("/arquivo/documento/<int:doc_id>/baixar")
def archive_download(doc_id: int):
    path = ARCHIVE.path_for(doc_id)
    if not path:
        abort(404)
    return send_file(path, mimetype="application/pdf", as_attachment=True, download_name=path.name, max_age=0)


@app.get("/arquivo/baixar.zip")
def archive_download_zip():
    year = request.args.get("year", type=int)
    month = request.args.get("month", type=int)
    q = request.args.get("q", "")
    company = request.args.get("company", "")
    exam_type = request.args.get("exam_type", "")
    receipt_filter = request.args.get("receipt_filter", "")
    unit_id = request.args.get("unit", "")
    raw_ids = request.args.get("ids", "")
    ids = []
    if raw_ids:
        for x in raw_ids.split(","):
            try:
                ids.append(int(x))
            except Exception:
                pass
    rows = ARCHIVE.filtered_rows(q=q, year=year, month=month, company_key=company, exam_type=exam_type, receipt_filter=receipt_filter, unit_id=unit_id, ids=ids or None)
    if not rows:
        return jsonify({"error": "Nenhum exame encontrado para este download."}), 404

    # O ZIP contém uma pasta raiz; empresas e tipos são organizados internamente conforme o conjunto filtrado.
    unique_companies = {str(r.get("company_key") or "") for r in rows}
    unique_periods = {(int(r.get("year", 0)), int(r.get("month", 0))) for r in rows}
    selected_company_name = ""
    if company and len(unique_companies) == 1:
        selected_company_name = str(rows[0].get("company_name") or "EMPRESA NAO IDENTIFICADA")
    elif len(unique_companies) == 1 and ids:
        selected_company_name = str(rows[0].get("company_name") or "EMPRESA NAO IDENTIFICADA")

    if year and month:
        period_label = f"{month:02d}-{year}"
    elif len(unique_periods) == 1:
        only_year, only_month = next(iter(unique_periods))
        period_label = f"{only_month:02d}-{only_year}"
    else:
        period_label = "TODOS OS PERIODOS"

    unit_label = ""
    if unit_id:
        unit = _unit_by_id(unit_id)
        unit_label = f"{unit['name']} - " if unit else ""
    if selected_company_name:
        folder_label = f"{unit_label}{period_label} - {selected_company_name}"
    elif ids:
        folder_label = f"{unit_label}{period_label} - EXAMES SELECIONADOS"
    else:
        folder_label = f"{unit_label}{period_label} - EXAMES FILTRADOS"
    safe_label = core.safe_component(folder_label, "EXAMES FILTRADOS", 150)
    mem = ARCHIVE.make_zip(rows, label=safe_label)
    return send_file(mem, mimetype="application/zip", as_attachment=True, download_name=f"{safe_label}.zip", max_age=0)



@app.get("/api/tipos-exames")
def api_exam_types():
    return jsonify({"types": core.get_custom_exam_types()})


@app.post("/api/tipos-exames")
def api_add_exam_type():
    data = request.get_json(silent=True) or request.form
    aliases_raw = data.get("aliases") or ""
    if isinstance(aliases_raw, str):
        aliases = [x.strip() for x in re.split(r"[,;\n]+", aliases_raw) if x.strip()]
    else:
        aliases = [str(x).strip() for x in aliases_raw if str(x).strip()]
    try:
        name = core.add_custom_exam_type(str(data.get("name") or ""), aliases=aliases)
        return jsonify({"ok": True, "name": name, "types": core.get_custom_exam_types()})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.delete("/api/tipos-exames/<path:name>")
def api_delete_exam_type(name: str):
    try:
        ok = core.delete_custom_exam_type(name)
        if not ok:
            return jsonify({"error": "Tipo padrão ou não encontrado não pode ser excluído."}), 400
        return jsonify({"ok": True, "types": core.get_custom_exam_types()})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/modelos")
def api_models():
    models = core.load_models()
    return jsonify({"models": [asdict(m) for m in models], "exam_types": list(core.get_exam_types())})


@app.post("/api/modelos")
def api_add_model():
    uploads = request.files.getlist("files") or request.files.getlist("file")
    uploads = [f for f in uploads if f and f.filename]
    exam_type = core.normalize_exam(request.form.get("exam_type") or "")
    label = (request.form.get("label") or "Modelo web").strip()
    if not uploads:
        return jsonify({"error": "Selecione um ou mais PDFs modelo."}), 400
    if any(Path(f.filename).suffix.lower() != ".pdf" for f in uploads):
        return jsonify({"error": "Todos os modelos precisam ser PDF."}), 400
    if not exam_type:
        return jsonify({"error": "Selecione o tipo de exame."}), 400
    folder = WORK_DIR / "model_uploads"
    folder.mkdir(parents=True, exist_ok=True)
    paths = []
    try:
        for f in uploads:
            path = _unique_upload_path(folder, f.filename)
            f.save(path)
            paths.append(path)
        added = core.add_models_from_pdfs(paths, exam_type, label, use_ocr=_load_config()["use_ocr"])
        return jsonify({"ok": True, "added": len(added), "files": len(paths)})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    finally:
        for path in paths:
            path.unlink(missing_ok=True)


@app.delete("/api/modelos/<model_id>")
def api_delete_model(model_id: str):
    return jsonify({"ok": core.delete_model(model_id)})


@app.get("/modelo-planilha")
def download_sheet_model():
    path = APP_DIR / "MODELO_FUNCIONARIOS.xlsx"
    return send_file(path, as_attachment=True, download_name="MODELO_FUNCIONARIOS.xlsx")


@app.get("/saude")
def health():
    return jsonify({"ok": True, "version": "5.0 WEB", "ocr": bool(core.locate_tesseract())})


def main() -> None:
    _cleanup_old_workspaces()
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8765"))
    if os.environ.get("EDGE_NO_BROWSER", "0") != "1" and host in {"127.0.0.1", "localhost"}:
        threading.Timer(1.2, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    try:
        from waitress import serve
        serve(app, host=host, port=port, threads=6)
    except ImportError:
        app.run(host=host, port=port, threaded=True, debug=False)


if __name__ == "__main__":
    main()
