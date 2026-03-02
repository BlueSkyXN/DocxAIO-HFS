#!/bin/sh
set -e

echo "=========================================="
echo "DocxAIO-HFS service starting"
echo "=========================================="
echo "PORT=${PORT:-8000}"
echo "MAX_FILE_SIZE_MB=${MAX_FILE_SIZE_MB:-120}"
echo "REQUEST_TIMEOUT_SECONDS=${REQUEST_TIMEOUT_SECONDS:-1200}"
echo "MAX_CONCURRENT_TASKS=${MAX_CONCURRENT_TASKS:-2}"
echo "TEMP_DIR=${TEMP_DIR:-/app/temp}"
echo "=========================================="

python --version
python -c "import fastapi, docx, openpyxl, matplotlib; print('Dependencies check: ok')"

exec uvicorn main:app \
    --host 0.0.0.0 \
    --port "${PORT:-8000}" \
    --workers "${WORKERS:-1}" \
    --log-level info \
    --access-log
