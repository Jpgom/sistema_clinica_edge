#!/usr/bin/env bash
set -o pipefail

echo "[build] Instalando dependências Python e tentando habilitar OCR..."

if command -v apt-get >/dev/null 2>&1; then
  echo "[build] Tentando instalar Tesseract OCR + idioma português via apt-get..."
  apt-get update && apt-get install -y tesseract-ocr tesseract-ocr-por libreoffice-writer fonts-liberation || \
    echo "[build] AVISO: apt-get não conseguiu instalar o Tesseract. O site sobe, mas PDFs escaneados/imagem não serão lidos por OCR."
else
  echo "[build] AVISO: apt-get não disponível neste ambiente."
fi

pip install -r requirements.txt

if command -v tesseract >/dev/null 2>&1; then
  echo "[build] OCR disponível:"
  tesseract --version | head -1
else
  echo "[build] AVISO FINAL: Tesseract não localizado. Configure o Build Command como: bash bin/render-build.sh"
fi
