# Correção - Envio periódicos / Comparecimento sem carregamento infinito

Ajuste aplicado em 17/09/2026.

## Problema
Após remover a opção "Considerar somente PERIÓDICO", a comparação passou a analisar todos os registros da planilha de controle. Em bases grandes, o código antigo fazia consultas repetidas ao banco para cada linha da planilha e mantinha o SQLite bloqueado durante toda a leitura do arquivo, podendo deixar o módulo Envio periódicos aparentemente carregando sem fim.

## Correção
- A planilha de controle agora é lida fora da transação de gravação do SQLite.
- Os convocados da competência são carregados uma única vez em memória.
- A comparação passa a usar dicionários rápidos por CNPJ+CPF, CPF e nome normalizado.
- O banco só é bloqueado no momento curto de gravar os resultados.
- A regra continua considerando todos os tipos de exame, sem filtro por PERIÓDICO.

## Resultado esperado
O comparecimento fica mais rápido e não bloqueia as demais telas do módulo durante a leitura da planilha.
