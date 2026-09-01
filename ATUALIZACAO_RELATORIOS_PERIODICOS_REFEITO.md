# Atualização: Relatório de Periódicos refeito

## Geração
- A função Relatórios usa sempre a coluna `ADMISSAO` para definir o mês.
- As colunas `VALIDADE`, `ULTIMO EXAME`, `VENCIMENTO` e `PERIODICO` não são usadas para o filtro mensal.
- Os dados exportados são padronizados em maiúsculo.
- No arquivo Relatório, duplicidades de `NOME + CARGO` são removidas.
- A Base do Mês reconhece `exames_obg` como Complementares.

## Comparação
- Nova etapa na página `/relatorios` para comparar o Relatório gerado pelo sistema com uma planilha de controle de ASO.
- O usuário envia o Relatório e a planilha de controle.
- O sistema lista as guias da planilha de controle para o usuário escolher.
- O relatório comparado recebe a coluna `DATA / TIPO DE EXAME / EMPRESA`.
- Quando houver mais de um registro do mesmo colaborador, todas as ocorrências são mantidas na mesma célula separadas por `|`.
- Nomes não encontrados permanecem no Relatório com a coluna adicional em branco.
