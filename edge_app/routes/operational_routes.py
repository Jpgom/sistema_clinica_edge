"""Rotas operacionais dos módulos da clínica."""
from edge_app import application as app_module

relatorios = app_module.relatorios
encaminhamentos = app_module.encaminhamentos
encaminhamento_especialista = app_module.encaminhamento_especialista
encaminhamento_especialista_gerar = app_module.encaminhamento_especialista_gerar
renumerador = app_module.renumerador
esocial = app_module.esocial
esocial_abas_base = app_module.esocial_abas_base
esocial_processar = app_module.esocial_processar
relatorios_async = app_module.relatorios_async
encaminhamentos_async = app_module.encaminhamentos_async
esocial_processar_async = app_module.esocial_processar_async
