# Correção Docker + Tesseract OCR no Render

O Render não permitiu instalar `tesseract-ocr` via `apt-get` no runtime Python nativo, retornando `Read-only file system`.

Esta versão adiciona:

- `Dockerfile` com Python 3.11.
- Instalação de `tesseract-ocr`, `tesseract-ocr-por` e `tesseract-ocr-eng` dentro da imagem Docker.
- `render.yaml` ajustado para `runtime: docker`.
- `.dockerignore` para reduzir arquivos desnecessários no build.
- Variáveis padrão para `/var/data`, mantendo dados persistentes.

No Render, altere o runtime do serviço de Python para Docker ou recrie o Web Service como Docker apontando para o mesmo repositório.
