FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY alimentar_maestro.py crm_drive.py crm_plus.py reporte_ventas_pdf.py app.py /app/
COPY templates/ /app/templates/

ENV PYTHONUNBUFFERED=1
ENV DATA_DIR=/data
ENV PORT=8000

EXPOSE $PORT

CMD uvicorn app:app --host 0.0.0.0 --port $PORT
