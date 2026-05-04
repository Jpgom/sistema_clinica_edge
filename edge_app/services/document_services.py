"""Serviços de geração/manipulação de documentos.

Centraliza o acesso aos processadores existentes para facilitar manutenção e testes.
"""
from edge_app import application as app_module

criar_relatorio = app_module.criar_relatorio
criar_base = app_module.criar_base
gerar_encaminhamentos = app_module.gerar_encaminhamentos
renumerar_documento = app_module.renumerar_documento
run_company_process = app_module.run_company_process
export_summary_excel = app_module.export_summary_excel
create_zip_from_folder = app_module.create_zip_from_folder
build_pdf = app_module.build_pdf
