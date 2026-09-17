# Correção - Envio periódicos carregando infinito

Ajustes aplicados:

- Redução do tempo de espera do SQLite para evitar tela carregando indefinidamente quando houver lock no banco.
- Recriação segura das pastas no disco persistente `/var/data/envio_periodicos`.
- Índices adicionados para acelerar painel, competências, convocações, anexos e logs de e-mail.
- Painel inicial refeito sem subconsultas pesadas por unidade.
- Rota de diagnóstico adicionada em `/envio-periodicos/health`.
- Tratamento de erro amigável no módulo para não deixar a tela branca/carregando.
- Migração `campaign_base_rows` deixou de ser refeita a cada deploy.
