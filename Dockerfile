FROM python:3.11-slim

ENV PORT=8000
ENV PYTHONUNBUFFERED=1
ENV TEMP_DIR=/app/temp
ENV MAX_FILE_SIZE_MB=120
ENV REQUEST_TIMEOUT_SECONDS=1200
ENV MAX_CONCURRENT_TASKS=2
ENV MPLCONFIGDIR=/app/.cache/matplotlib

RUN apt-get update && apt-get install -y --no-install-recommends \
    fontconfig \
    fonts-noto-cjk \
    fonts-wqy-zenhei \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .
COPY docx_allinone.py .
COPY entrypoint.sh .
COPY templates/ ./templates/
COPY static/ ./static/

RUN chmod +x /app/entrypoint.sh \
    && mkdir -p /app/temp /app/.cache/matplotlib \
    && chmod -R 777 /app/temp /app/.cache/matplotlib

EXPOSE 8000

ENTRYPOINT ["/app/entrypoint.sh"]
