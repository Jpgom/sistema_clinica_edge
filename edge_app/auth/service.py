"""Serviços de autenticação e autorização."""
from edge_app import application as app_module

init_auth_db = app_module.init_auth_db
auth_current_user = app_module.auth_current_user
auth_is_logged_in = app_module.auth_is_logged_in
auth_is_admin = app_module.auth_is_admin
has_role = app_module.has_role
auth_create_user = app_module.auth_create_user
auth_update_user = app_module.auth_update_user
auth_delete_user = app_module.auth_delete_user
auth_change_own_password = app_module.auth_change_own_password
audit_log = app_module.audit_log
