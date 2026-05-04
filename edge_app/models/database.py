"""Acesso a banco e helpers de compatibilidade SQLite/PostgreSQL."""
from edge_app import application as app_module

auth_get_conn = app_module.auth_get_conn
db_param = app_module.db_param
db_id_type = app_module.db_id_type
db_fetchone = app_module.db_fetchone
db_fetchall = app_module.db_fetchall
db_execute = app_module.db_execute
row_get = app_module.row_get
