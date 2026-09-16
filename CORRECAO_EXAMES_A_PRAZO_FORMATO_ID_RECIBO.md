# Correção - Exames a Prazo: ID TRANSAÇÃO e RECIBO

Ajuste aplicado na função Exames a Prazo.

## Problema
As colunas **ID TRANSAÇÃO** e **RECIBO** herdavam o estilo da planilha/modelo de saída, podendo aparecer no Excel como valor monetário/contábil.

## Correção
- A leitura agora captura o formato original das células da planilha-base.
- Na geração, as colunas **ID TRANSAÇÃO** e **RECIBO** usam o formato original da base quando ele é confiável.
- Se o formato da base estiver ausente ou vier como moeda, o sistema força a exibição como identificador, usando formato Geral/Texto.
- A coluna **VALOR** permanece com formato monetário.

## Resultado esperado
Os dados de **ID TRANSAÇÃO** e **RECIBO** passam a sair como na base, sem serem exibidos como dinheiro.
