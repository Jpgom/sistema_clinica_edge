FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    TESSDATA_PREFIX=/usr/share/tesseract-ocr/5/tessdata \
    OMP_THREAD_LIMIT=1 \
    RENDER_DISK_PATH=/var/data \
    ENVIO_PERIODICOS_DATA_DIR=/var/data/envio_periodicos \
    SEPARAR_EXAMES_DATA_DIR=/var/data/separar_exames

# Dependências de sistema necessárias para OCR e processamento de PDF/imagem.
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        tesseract-ocr \
        tesseract-ocr-por \
        tesseract-ocr-eng \
        poppler-utils \
        libgl1 \
        libglib2.0-0 \
        curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt ./
RUN pip install --upgrade pip setuptools wheel \
    && pip install -r requirements.txt

COPY . .

RUN mkdir -p /var/data/envio_periodicos /var/data/separar_exames

EXPOSE 10000

CMD ["sh", "-c", "gunicorn app:app --bind 0.0.0.0:${PORT:-10000} --workers 1 --timeout 300 --graceful-timeout 300 --threads 4"]
