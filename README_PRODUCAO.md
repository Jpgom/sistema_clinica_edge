# Sistema EDGE — Versão refatorada para produção

## O que mudou nesta versão

- `app.py` agora é apenas o ponto de entrada WSGI.
- A aplicação principal foi movida para `edge_app/application.py`.
- Foi criada uma fila assíncrona leve em `edge_app/async_jobs.py`.
- Os fluxos pesados de Relatórios, Encaminhamentos e E-SOCIAL agora usam páginas de processamento com progresso e download posterior.
- Foi adicionada a página `templates/job_status.html`.
- Foram adicionados estilos de barra de progresso no CSS.
- Os formulários principais foram redirecionados para endpoints assíncronos.

## Como rodar localmente

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Acesse:

```text
http://127.0.0.1:5000
```

## Variáveis importantes no Render

```text
SECRET_KEY=gere_uma_chave_segura
DATABASE_URL=sua_url_postgresql
SESSION_COOKIE_SECURE=1
MAX_CONTENT_LENGTH=78643200
JOB_WORKERS=2
JOB_TTL_HOURS=12
```

## Recomendação de produção

Para uso real da clínica:

- Render Web Service Standard
- PostgreSQL Starter ou superior
- `SECRET_KEY` fixa em variável de ambiente
- Disco persistente ou armazenamento externo para arquivos temporários importantes

## Observação

A fila assíncrona desta versão usa threads internas. Ela é adequada para uso pequeno/médio. Para alto volume, substitua por Celery + Redis mantendo os mesmos endpoints de status/download.
