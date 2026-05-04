# Deploy no Render com login empresarial e PostgreSQL

Este projeto é Flask e deve ser publicado como Web Service Python.

## Build Command
pip install -r requirements.txt

## Start Command
gunicorn app:app

## Variáveis obrigatórias
SECRET_KEY=uma_chave_grande_e_segura
DATABASE_URL=cole_a_internal_database_url_do_postgresql
FLASK_ENV=production
SESSION_COOKIE_SECURE=1
SESSION_HOURS=8

## Variáveis opcionais para criar admin automático
ADMIN_USERNAME=admin
ADMIN_PASSWORD=sua_senha_forte_com_letras_e_numeros
ADMIN_NAME=Administrador

Se ADMIN_USERNAME e ADMIN_PASSWORD forem definidos, o sistema cria o admin automaticamente na primeira inicialização, caso ele ainda não exista.

Depois do primeiro login, é recomendado remover ADMIN_PASSWORD do ambiente ou trocar a senha pela tela Minha conta.


Arquitetura: consulte ARQUITETURA_MODULAR.md para a estrutura edge_app/routes, services, workers, models, auth, utils, templates e static.
