"""Ponto de entrada WSGI do Sistema EDGE.

A aplicação principal foi movida para `edge_app.application` para deixar este
arquivo pequeno, previsível e compatível com Render/Gunicorn.
"""
from edge_app.application import app

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
