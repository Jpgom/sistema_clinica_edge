# Correção - Separar exames não encontrava todos os exames

Ajustes aplicados:

- Tipos de exames cadastrados pelo usuário agora têm peso alto na identificação por texto.
- Apelidos/sinônimos dos exames também viram assinaturas de reconhecimento.
- A leitura da planilha não quebra mais nomes de exames com barra, como RAIO-X / TÓRAX.
- A varredura de segurança usa modelo visual quando houver modelo cadastrado para exame faltante.
- Exames personalizados com modelo cadastrado aceitam confiança menor quando funcionário/CPF/CNPJ conferem fortemente.
- Incluído script de build para tentar instalar Tesseract OCR no Render, necessário para PDFs escaneados.

Observação: PDFs digitais com texto interno funcionam sem OCR. PDFs escaneados/imagem dependem do binário `tesseract` no ambiente.
