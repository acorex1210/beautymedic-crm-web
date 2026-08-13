# Web de reportes Derma Essenza

App web para generar los reportes PDF de ventas por campaña/CRM y sincronizar
el maestro `BD DATA.xlsx` desde Google Drive, igual que los scripts de terminal
(`reporte_ventas_pdf.py` y `alimentar_maestro.py`).

## Ejecutar local

```bash
cd reportes_web
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
# (opcional) apuntar al maestro y credenciales locales
export MAESTRO_PATH="$HOME/Downloads/BD DATA.xlsx"
export CREDENCIALES="$HOME/credenciales-bm.json"
uvicorn app:app --reload --port 8000
```

Abrir http://localhost:8000

## Desplegar en la nube

La app es un solo contenedor (FastAPI + uvicorn) autocontenido en este
directorio. Requiere dos volúmenes/persistencia:

| Variable            | Ejemplo                              | Uso |
|---------------------|--------------------------------------|-----|
| `DATA_DIR`          | `/data` (volumen persistente)        | Reportes PDF y backups |
| `MAESTRO_PATH`      | `/data/BD DATA.xlsx`                 | Maestro (se sube/descarga por la web) |
| `TMP_DIR`           | `/data/tmp`                          | Descargas de Drive y simulaciones |
| `CREDENCIALES`      | `/data/credenciales.json`            | JSON de cuenta de servicio de Drive |
| `GDRIVE_CREDENTIALS_JSON` | (contenido del JSON)           | Alternativa: se escribe automáticamente |
| `AGENDADOS_FID`     | `12fWJpIBpr3GH7Yj57iyyndm_m37rr7V2` | ID del archivo AGENDADOS en Drive |
| `VENTA_FID`         | `1LHtZk0vAGgnyOsODwU6f4LvtUoQxWNis` | ID del archivo VENTA DIARIA en Drive |

Pasos típicos:

1. Montar un volumen persistente en `DATA_DIR` (p. ej. Railway volume `/data`
   o Render disk `/data`).
2. Subir las credenciales de Drive: ya sea por variable
   `GDRIVE_CREDENTIALS_JSON` (contenido completo del JSON) o copiando el
   archivo `credenciales.json` al volumen.
3. Subir el maestro `BD DATA.xlsx` una vez desde la pestaña "Maestro" de la
   web (se guarda en `MAESTRO_PATH`).
4. Definir `AGENDADOS_FID`, `VENTA_FID`, `MAESTRO_PATH` y `TMP_DIR`.

### Despliegue con Railway CLI (desde este directorio)

```bash
# crear el servicio y el volumen persistente
railway add --service reportes
railway volume add --mount-path /data        # se auto-adjunta al servicio
# variables (ver tabla de arriba)
railway variable set "MAESTRO_PATH=/data/BD DATA.xlsx" "TMP_DIR=/data/tmp" \
  "CREDENCIALES=/data/credenciales.json" "AGENDADOS_FID=..." "VENTA_FID=..." \
  "PORT=8000" --service reportes --skip-deploys
cat ~/credenciales-bm.json | railway variable set GDRIVE_CREDENTIALS_JSON \
  --stdin --service reportes          # credenciales (la app las escribe al volumen)
# desplegar y obtener URL
railway up --service reportes --detach
railway domain --service reportes
```

Tras el primer deploy, sube el maestro una vez desde la pestaña "Maestro" de la
web (o `curl -F "file=@BD DATA.xlsx" <url>/api/maestro/upload`). Los cambios
posteriores de código se despliegan con `railway up` desde este directorio.

Los archivos `alimentar_maestro.py` y `reporte_ventas_pdf.py` viven aquí; en la
raíz del repo quedan wrappers que ejecutan el mismo código (CLI sin cambios).

## API

| Método | Ruta                          | Descripción |
|--------|-------------------------------|-------------|
| GET    | `/`                           | Página web |
| GET    | `/api/estado`                 | Estado: credenciales, maestro, reportes |
| POST   | `/api/reporte`                | Genera PDF `{mes, anio, desde, hasta, fuente}` |
| GET    | `/api/reportes`               | Lista de PDFs generados |
| GET    | `/api/reporte/download/{f}`   | Descarga un PDF |
| DELETE | `/api/reporte/{f}`            | Borra un PDF |
| POST   | `/api/sync`                   | Sincroniza `{aplicar: bool}` (aplicar=false = revisar) |
| POST   | `/api/maestro/upload`         | Sube el maestro (multipart `file`) |
| GET    | `/api/maestro/download`       | Descarga el maestro |
