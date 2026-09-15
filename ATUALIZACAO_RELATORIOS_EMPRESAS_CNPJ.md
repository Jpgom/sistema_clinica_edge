# Atualização - Relatórios com cadastro oficial de empresas por CNPJ

Alterações adicionadas:

1. Nova área em Relatórios para importar uma planilha de cadastro com EMPRESA e CNPJ.
2. O cadastro fica salvo em `DATA_DIR/relatorios_empresas_cnpj.json` ou no disco persistente do Render, quando configurado.
3. A geração do Relatório de Periódicos continua usando sempre a coluna ADMISSAO para filtrar o mês.
4. Quando a planilha de convocação possuir a coluna CNPJ, o sistema procura esse CNPJ no cadastro oficial.
5. No arquivo Relatorio_MES.xlsx, o título do bloco passa a sair como `EMPRESA - CNPJ`.
6. No arquivo Base_do_Mes_MES.xlsx, as colunas são `EMPRESA`, `CNPJ`, `NOME`, `CARGO` e `COMPLEMENTARES`.
7. O CNPJ sai formatado como `00.000.000/0000-00` sempre que houver 14 dígitos ou quando o Excel remover zeros à esquerda.
8. Caso o CNPJ não esteja cadastrado, o sistema usa o nome que veio na planilha de convocação e mantém o CNPJ formatado.

Não é necessário criar nova variável de ambiente.
