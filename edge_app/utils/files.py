"""Utilitários de arquivos, uploads e nomes seguros."""
from edge_app import application as app_module

normalize_text = app_module.normalize_text
sanitize_filename = app_module.sanitize_filename
validate_uploaded_file = app_module.validate_uploaded_file
unique_path = app_module.unique_path
safe_extract_zip = app_module.safe_extract_zip
limpar_nome_arquivo = app_module.limpar_nome_arquivo
limpar_nome_pasta_arquivo = app_module.limpar_nome_pasta_arquivo
