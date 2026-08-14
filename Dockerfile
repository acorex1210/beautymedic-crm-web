FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY *.py /app/
COPY templates/ /app/templates/
COPY static/ /app/static/

# Crear directorio persistente para backups, reportes y datos
RUN mkdir -p /app/data

ENV PYTHONUNBUFFERED=1
ENV TMP_DIR=/tmp
ENV DATA_DIR=/app/data
ENV PORT=8080

EXPOSE $PORT

CMD uvicorn app:app --host 0.0.0.0 --port $PORT