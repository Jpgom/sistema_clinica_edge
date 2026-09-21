#!/usr/bin/env bash
set -o pipefail

# Build usado no Render. O pytesseract é apenas o wrapper Python; para ler PDF
# escaneado/imagem o binário do Tesseract precisa existir no ambiente.
if command -v apt-get >/dev/null 2>&1; then
  echo "[build] Tentando instalar Tesseract OCR e idioma português..."
  (apt-get update && apt-get install -y tesseract-ocr tesseract-ocr-por) || \
    echo "[build] AVISO: não foi possível instalar Tesseract via apt-get. PDFs digitais continuam funcionando; PDFs escaneados podem exigir Docker ou instalação do Tesseract no ambiente."
fi

pip install -r requirements.txt
