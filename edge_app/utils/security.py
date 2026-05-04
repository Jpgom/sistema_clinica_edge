"""Utilitários de segurança de sessão, CSRF e senhas."""
from edge_app import application as app_module

validate_password_strength = app_module.validate_password_strength
generate_csrf_token = app_module.generate_csrf_token
validate_csrf = app_module.validate_csrf
apply_security_headers = app_module.apply_security_headers
