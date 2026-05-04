# Arquitetura modular do Sistema EDGE

A estrutura do projeto foi reorganizada para o padrão profissional solicitado:

```text
edge_app/
  routes/      # agrupamento dos endpoints por domínio
  services/    # serviços de processamento de documentos e regras de negócio
  workers/     # fila assíncrona e jobs pesados
  models/      # acesso a banco e modelos auxiliares
  auth/        # autenticação, autorização e auditoria
  utils/       # arquivos, segurança e utilidades compartilhadas
  templates/   # templates Jinja2 dentro do pacote da aplicação
  static/      # CSS/JS/assets dentro do pacote da aplicação
```

## Observação técnica

Para reduzir risco de regressão antes do deploy, os endpoints continuam registrados em
`edge_app/application.py`, mas os domínios agora existem como módulos separados com
facades oficiais. Isso permite evoluir de forma segura para Blueprints completos sem
quebrar URLs, templates, jobs, autenticação ou deploy no Render.

## Compatibilidade Render

O entrypoint continua sendo:

```bash
gunicorn app:app
```

Nenhuma variável de ambiente foi removida.
