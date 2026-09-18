# Atualização - Separar exames

Adicionada a função **Separar exames** ao site principal em `/separar-exames/`.

A função foi integrada a partir do modelo `EDGE_EXTRATOR_v5.0_CONFIABILIDADE_FINAL`, mantendo o fluxo original:

- upload da planilha de funcionários/exames;
- upload de PDFs de exames;
- processamento em lote;
- revisão de pendências/duplicidades;
- download do ZIP separado e do relatório de conferência;
- arquivo de exames processados.

Em produção, os dados persistentes ficam em `SEPARAR_EXAMES_DATA_DIR`. Se o Render tiver `RENDER_DISK_PATH=/var/data`, o sistema usa automaticamente `/var/data/separar_exames`.
