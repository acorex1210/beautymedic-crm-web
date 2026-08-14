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

### Web local con URL pública fija (sin hosting de pago)

Para usar la web de Derma Essenza desde cualquier lugar sin pagar hosting, la
app corre en esta Mac con uvicorn y se publica con Tailscale Funnel (gratis):

```bash
bash iniciar_derma.sh
```

Ese script levanta el servidor local (puerto 8011) y el túnel público, y
muestra las URLs:

| Acceso | URL |
|--------|-----|
| Local | http://127.0.0.1:8011 |
| Pública (fija) | https://macbook-neo-de-andre.tailab4d2b.ts.net |

Requisitos (una sola vez):
- `run_derma_local.sh` usa las credenciales de Drive en `~/credenciales-derma.json`,
  el maestro `BD DATA DERMA ESSENZA.xlsx` en `~/Downloads` y los `FID` de Derma
  Essenza (AGENDADOS `1So_1Fh744c3K9kss2oA1twjBLJpgrSxZCu2lqhWpqJM`,
  VENTA `1TDM7ZFV6Jdsqc6i4CadNkwPQNdrIBhu7`).
- Tailscale app instalada, iniciada sesión y Funnel habilitado en
  https://login.tailscale.com/f/funnel para el nodo.

Detener la web pública:
```bash
tailscale funnel --https=443 off
kill $(pgrep -f 'uvicorn app:app')
```

> La web pública depende de que la Mac esté encendida y con Tailscale activo.
> Para una opción "siempre encendida" sin depender de la Mac, usar Cloud Run
> (gratuito, sección "Desplegar en la nube").

## Desplegar en la nube

La app es un solo contenedor (FastAPI + uvicorn) autocontenido en este
directorio. Requiere dos volúmenes/persistencia:

| Variable            | Ejemplo                              | Uso |
|---------------------|--------------------------------------|-----|
| `DATA_DIR`          | `/data` (volumen persistente)        | Reportes PDF y backups |
| `MAESTRO_PATH`      | `/data/BD DATA.xlsx`                 | Maestro local (se sube/descarga por la web) |
| `MAESTRO_FID`       | `1aBc...` (ID en Drive)              | Maestro en Google Drive (host sin disco: Cloud Run) |
| `TMP_DIR`           | `/data/tmp`                          | Descargas de Drive y simulaciones |
| `CREDENCIALES`      | `/data/credenciales.json`            | JSON de cuenta de servicio de Drive |
| `GDRIVE_CREDENTIALS_JSON` | (contenido del JSON)           | Alternativa: se escribe automáticamente |
| `AGENDADOS_FID`     | `12fWJpIBpr3GH7Yj57iyyndm_m37rr7V2` | ID del archivo AGENDADOS en Drive |
| `VENTA_FID`         | `1LHtZk0vAGgnyOsODwU6f4LvtUoQxWNis` | ID del archivo VENTA DIARIA en Drive |

> `MAESTRO_FID` hace al maestro portable: si está definido, el maestro se lee y
> escribe desde Google Drive (ideal para hosts sin disco persistente como Cloud
> Run). Sin `MAESTRO_FID` se usa `MAESTRO_PATH` local, como siempre en Railway.

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

### Despliegue en Google Cloud Run (gratuito, sin disco)

Cloud Run no tiene disco persistente, por eso el maestro y las credenciales van
en Drive: basta definir `MAESTRO_FID` (además de `AGENDADOS_FID` y `VENTA_FID`).

1. Crea un proyecto en https://console.cloud.google.com (requiere vincular una
   tarjeta; el free tier no cobra nada). Anota el **ID** del proyecto.
2. Abre **Cloud Shell** (ícono `>_` arriba a la derecha) y sube este directorio
   `reportes_web` (botón "Subir archivos" / arrastrar un `.zip` y descomprimir):
   ```bash
   unzip reportes_web_cloudrun.zip && cd reportes_web
   ```
3. Ejecuta (el script te pedirá el ID del proyecto):
   ```bash
   bash deploy_cloudrun.sh
   ```
   Al final muestra la URL de la página (terminada en `run.app`).

La app se escala a cero cuando nadie la usa, así que no genera costo; se
"despierta" sola con el primer clic (unos segundos).

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
