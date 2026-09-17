# Correção - Envio periódicos carregamento infinito

Ajustes aplicados:

1. O painel inicial do Envio periódicos foi deixado mais leve para não fazer consultas pesadas em todas as convocações ao abrir.
2. A tela de unidades também foi otimizada para listar competências sem travar por contagens grandes.
3. Jobs antigos de envio que ficaram em estado QUEUED/RUNNING passam para CONFERÊNCIA, sem apagar competências, quando estiverem antigos/travados.
4. O worker de envio de e-mail deixa de iniciar automaticamente no boot do site e inicia somente quando o usuário solicitar novo envio.
5. Adicionada rota `/envio-periodicos/repair` para destravar fila antiga sem apagar dados.
6. Removida abertura automática do modal de andamento quando já existe job ativo antigo, evitando aparência de carregamento infinito.
7. Adicionados índices extras para acelerar consultas em email_logs e jobs.

Nenhuma competência é apagada por esta correção.
