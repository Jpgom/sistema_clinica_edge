# Correção Separar exames - OCR obrigatório

- Adicionada verificação clara de OCR no painel.
- Adicionada rota `/separar-exames/api/ocr-status` para diagnóstico.
- Se os PDFs forem escaneados/imagem e o OCR não estiver instalado, o sistema bloqueia o processamento com mensagem direta em vez de gerar todos como não encontrados.
- Build Command recomendado no Render: `bash bin/render-build.sh`.
