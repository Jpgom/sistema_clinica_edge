# Correção - Exames a Prazo / FROTAS

Ajustes aplicados:

- Correção da leitura de guias sem cabeçalho na primeira linha, como `JANEIRO.2025 OK`.
- Quando a guia não possui cabeçalho detectável, o sistema passa a usar o layout operacional padrão:
  - A: Funcionário
  - C: Recibo
  - D: Valor
  - E: Nº do exame
  - F: Tipo de exame
  - G: Função
  - H: ID transação
  - I: Depositante
  - J: Data
  - K: Status
  - L: Setor/empresa
- Mantida compatibilidade com guias antigas de 10 colunas.
- Limite padrão de upload do sistema aumentado para 250 MB.
- Limite individual de base XLSX do Exames a Prazo aumentado para 150 MB por padrão, ajustável por `EXAMES_A_PRAZO_MAX_SINGLE_XLSX_MB`.
- ZIPs de entrada com CNPJs passam a aceitar até 600 MB descompactados por padrão, ajustável por `EXAMES_A_PRAZO_MAX_ZIP_MB`.
- O ZIP de saída agora inclui `RESUMO_EXAMES_A_PRAZO.txt` com competências selecionadas, CNPJs solicitados e quantidade encontrada por CNPJ.
- A tela recebeu botões rápidos para selecionar 2025, 2026, todas as competências ou limpar seleção.

Teste realizado com:

- `PLANILHA E-SOCIAL - 2025.2026 - BELÉM (3).xlsx`
- `FROTAS.xlsx`
- Todas as competências de 2025

Resultado do teste: 139 registros encontrados para os CNPJs informados.
