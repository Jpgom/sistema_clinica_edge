"""Interface HTTP dos recibos, compartilhando a sessão e proteções do EDGE."""

from __future__ import annotations

import base64
import json
import os
import tempfile
from io import BytesIO
from pathlib import Path

from flask import Blueprint, current_app, jsonify, render_template, request, send_file, session
from werkzeug.exceptions import RequestEntityTooLarge

from .core import (
    build_layout,
    load_default_config,
    normalize_config,
    render_pdf,
    render_preview_png,
)


recibos = Blueprint(
    "recibos",
    __name__,
    url_prefix="/recibos",
    template_folder="templates",
    static_folder="static",
)

_MAX_REQUEST_BYTES = 128 * 1024
_MAX_CONFIG_BYTES = 64 * 1024


def _user_config_path() -> Path:
    """Derive o caminho apenas da sessão validada pelo EDGE, nunca do pedido."""
    try:
        user_id = int(session["user_id"])
    except (KeyError, TypeError, ValueError) as exc:
        raise ValueError("Sessão inválida. Entre novamente no sistema.") from exc
    if user_id <= 0:
        raise ValueError("Sessão inválida. Entre novamente no sistema.")
    data_dir = Path(current_app.config["RECIBOS_DATA_DIR"])
    return data_dir / f"calibracao_usuario_{user_id}.json"


def _config_size(config: dict) -> int:
    try:
        return len(json.dumps(config, ensure_ascii=False, allow_nan=False).encode("utf-8"))
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError("A calibração contém valores inválidos.") from exc


def _validate_config(value: object) -> dict:
    if not isinstance(value, dict):
        raise ValueError("A calibração deve ser um objeto JSON.")
    if _config_size(value) > _MAX_CONFIG_BYTES:
        raise ValueError("A calibração excede o tamanho permitido.")
    config = normalize_config(value)
    if not isinstance(config, dict):
        raise ValueError("A calibração é inválida.")
    if _config_size(config) > _MAX_CONFIG_BYTES:
        raise ValueError("A calibração excede o tamanho permitido.")
    return config


def _load_config() -> dict:
    default = _validate_config(load_default_config())
    path = _user_config_path()
    try:
        if path.stat().st_size > _MAX_CONFIG_BYTES:
            raise ValueError("Calibração salva acima do limite permitido.")
        with path.open("r", encoding="utf-8") as handle:
            return _validate_config(json.load(handle))
    except FileNotFoundError:
        return default
    except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
        current_app.logger.warning("Calibração de recibo inválida em %s: %s", path, exc)
        return default


def _save_config(config: dict) -> None:
    path = _user_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    contents = json.dumps(config, ensure_ascii=False, allow_nan=False, indent=2) + "\n"
    if len(contents.encode("utf-8")) > _MAX_CONFIG_BYTES:
        raise ValueError("A calibração excede o tamanho permitido.")
    temp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=path.parent, prefix=".calibracao-",
            suffix=".tmp", delete=False,
        ) as handle:
            temp_path = handle.name
            handle.write(contents)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, path)
    finally:
        if temp_path and os.path.exists(temp_path):
            os.unlink(temp_path)


def _json_body() -> dict:
    if not request.is_json:
        raise ValueError("Envie os dados em formato JSON.")
    if request.content_length is not None and request.content_length > _MAX_REQUEST_BYTES:
        raise RequestEntityTooLarge()
    raw = request.get_data(cache=True)
    if len(raw) > _MAX_REQUEST_BYTES:
        raise RequestEntityTooLarge()
    body = request.get_json(silent=True)
    if not isinstance(body, dict):
        raise ValueError("O JSON enviado é inválido.")
    return body


def _payload_and_config() -> tuple[dict, dict]:
    body = _json_body()
    payload = body.get("payload")
    if not isinstance(payload, dict):
        raise ValueError("Os dados do recibo são inválidos.")
    config_value = body.get("config")
    config = _load_config() if config_value is None else _validate_config(config_value)
    return payload, config


@recibos.errorhandler(ValueError)
def _invalid_input(error: ValueError):
    return jsonify(error=str(error)), 400


@recibos.errorhandler(RequestEntityTooLarge)
def _request_too_large(_error: RequestEntityTooLarge):
    return jsonify(error="Os dados enviados excedem o limite de 128 KB."), 413


@recibos.get("/")
def index():
    # A permissão é conferida no banco pelo EDGE. O papel na sessão pode estar desatualizado.
    from edge_app.application import is_readonly_user

    return render_template(
        "recibo/index.html", title="Recibos", config=_load_config(),
        can_edit=not is_readonly_user(),
    )


@recibos.post("/api/preview")
def api_preview():
    payload, config = _payload_and_config()
    layout = build_layout(payload, config)
    png = render_preview_png(payload, config)
    response = jsonify(image="data:image/png;base64," + base64.b64encode(png).decode("ascii"), layout=layout)
    response.headers["Cache-Control"] = "no-store"
    return response


@recibos.post("/api/pdf")
def api_pdf():
    body = _json_body()
    payload = body.get("payload")
    if not isinstance(payload, dict):
        raise ValueError("Os dados do recibo são inválidos.")
    config_value = body.get("config")
    config = _load_config() if config_value is None else _validate_config(config_value)
    with_model = body.get("with_model", False)
    if not isinstance(with_model, bool):
        raise ValueError("A opção de modelo é inválida.")
    pdf = render_pdf(payload, config, with_model=with_model)
    response = send_file(
        BytesIO(pdf), mimetype="application/pdf", as_attachment=True,
        download_name="recibo_conferencia.pdf" if with_model else "recibo_dados.pdf",
    )
    response.headers["Cache-Control"] = "no-store"
    return response


@recibos.post("/api/config")
def api_config():
    body = _json_body()
    config = _validate_config(body.get("config"))
    _save_config(config)
    response = jsonify(ok=True, config=config)
    response.headers["Cache-Control"] = "no-store"
    return response
