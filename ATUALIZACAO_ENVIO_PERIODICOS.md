# Atualização - ENVIO PERIÓDICOS

Esta versão integra o módulo EDGE Convocações Periódicos V5.1 ao site principal no caminho `/envio-periodicos/`.

## O que foi incluído
- Nova aba **ENVIO PERIÓDICOS** no menu e na tela inicial.
- Cadastro/importação de empresas por CNPJ, e-mail, CC e responsável.
- Cadastro de unidades BELÉM e MACAPÁ.
- Criação de competências por mês/ano.
- Importação de planilhas de convocação.
- Deduplicação de colaboradores por empresa.
- Geração de planilha base para encaminhamentos somente por botão.
- Controle de comparecimento para reenvio/lembrete de pendentes.
- Envio de e-mails por Gmail/SMTP com fila persistente, progresso em tela e proteção contra duplo clique.
- Exclusão definitiva de competência com confirmação.

## Render
Recomendado usar um disco persistente e configurar `ENVIO_PERIODICOS_DATA_DIR=/var/data/envio_periodicos`.
Se já existir `RENDER_DISK_PATH=/var/data`, o sistema usa automaticamente `/var/data/envio_periodicos`.

Mantenha o start command com apenas 1 worker. Exemplo:
`gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 300 --graceful-timeout 300`

