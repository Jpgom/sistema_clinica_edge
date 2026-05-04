"""Rotas de autenticação e gestão de usuários.

As rotas continuam registradas em edge_app.application nesta versão para preservar
compatibilidade de endpoints, templates e regras de sessão. Este módulo documenta
o domínio e serve como ponto oficial para a próxima extração completa por Blueprint.

Endpoints relacionados:
- /primeiro-acesso
- /login
- /logout
- /minha-conta
- /auditoria
- /usuarios
"""

from edge_app import application as app_module

setup_admin = app_module.setup_admin
login = app_module.login
logout = app_module.logout
minha_conta = app_module.minha_conta
auditoria = app_module.auditoria
usuarios = app_module.usuarios
usuarios_criar = app_module.usuarios_criar
usuarios_editar = app_module.usuarios_editar
usuarios_excluir = app_module.usuarios_excluir
