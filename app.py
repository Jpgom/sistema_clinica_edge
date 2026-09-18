"""Ponto de entrada WSGI do Sistema EDGE integrado ao módulo PGR/SST.

- O sistema principal continua respondendo em `/`.
- O módulo PGR/SST fica montado em `/pgr`.
- O módulo Envio Periódicos fica montado em `/envio-periodicos`.
- A função Separar exames fica montada em `/separar-exames`.
"""
from __future__ import annotations

import os

from werkzeug.middleware.dispatcher import DispatcherMiddleware
from werkzeug.serving import run_simple

from edge_app.application import app as edge_app

# Configuração padrão do módulo Envio Periódicos quando está integrado ao site principal.
# O módulo usa SQLite e arquivos internos; em produção, a pasta abaixo deve ficar em disco persistente.
_persist_root = os.environ.get("RENDER_DISK_PATH") or os.environ.get("DATA_DIR")
if _persist_root and "ENVIO_PERIODICOS_DATA_DIR" not in os.environ:
    os.environ["ENVIO_PERIODICOS_DATA_DIR"] = os.path.join(_persist_root, "envio_periodicos")
if _persist_root and "SEPARAR_EXAMES_DATA_DIR" not in os.environ:
    os.environ["SEPARAR_EXAMES_DATA_DIR"] = os.path.join(_persist_root, "separar_exames")
os.environ.setdefault("EDGE_LOCAL_AUTH", "0")

from pgr_app.app import app as pgr_app
from envio_periodicos_app.app import app as envio_periodicos_app
from separar_exames_app.app import app as separar_exames_app

# Garante que os módulos leiam o mesmo cookie de sessão/login.
for _mounted_app in (pgr_app, envio_periodicos_app, separar_exames_app):
    _mounted_app.secret_key = edge_app.secret_key
    for key in ("SESSION_COOKIE_HTTPONLY", "SESSION_COOKIE_SAMESITE", "SESSION_COOKIE_SECURE"):
        _mounted_app.config[key] = edge_app.config.get(key)

# `app` é o objeto usado pelo Gunicorn no Render: `gunicorn app:app`.
app = DispatcherMiddleware(edge_app, {
    "/pgr": pgr_app,
    "/envio-periodicos": envio_periodicos_app,
    "/separar-exames": separar_exames_app,
})

# Alias opcional para ferramentas que procuram `application`.
application = app

if __name__ == "__main__":
    run_simple("0.0.0.0", int(os.environ.get("PORT", "5000")), app, use_debugger=True)
