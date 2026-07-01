"""Ponto de entrada WSGI do Sistema EDGE integrado ao módulo PGR/SST.

- O sistema principal continua respondendo em `/`.
- O módulo PGR/SST fica montado em `/pgr`.
"""
from __future__ import annotations

import os

from werkzeug.middleware.dispatcher import DispatcherMiddleware
from werkzeug.serving import run_simple

from edge_app.application import app as edge_app
from pgr_app.app import app as pgr_app

# Garante que os dois módulos leiam o mesmo cookie de sessão/login.
pgr_app.secret_key = edge_app.secret_key
for key in ("SESSION_COOKIE_HTTPONLY", "SESSION_COOKIE_SAMESITE", "SESSION_COOKIE_SECURE"):
    pgr_app.config[key] = edge_app.config.get(key)

# `app` é o objeto usado pelo Gunicorn no Render: `gunicorn app:app`.
app = DispatcherMiddleware(edge_app, {
    "/pgr": pgr_app,
})

# Alias opcional para ferramentas que procuram `application`.
application = app

if __name__ == "__main__":
    run_simple("0.0.0.0", int(os.environ.get("PORT", "5000")), app, use_debugger=True)
