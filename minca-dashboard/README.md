# Minca Electric — Dashboard Operativo

Dashboard interactivo que se actualiza automáticamente desde Google Drive.

## Variables de entorno requeridas en Vercel

```
GOOGLE_SERVICE_ACCOUNT={"type":"service_account","project_id":"...","private_key_id":"...","private_key":"-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n","client_email":"...@....gserviceaccount.com",...}
```

## Setup Google Service Account

1. Ve a https://console.cloud.google.com
2. Crea un proyecto → Activa **Google Drive API**
3. IAM & Admin → Service Accounts → Crear
4. Descarga el JSON de credenciales
5. Comparte el archivo de Drive con el email de la service account (Lector)
6. Pega el contenido del JSON como valor de `GOOGLE_SERVICE_ACCOUNT` en Vercel

## Deploy

```bash
git push origin main
```
Vercel detecta el push y despliega automáticamente.

## Actualización automática

El dashboard lee el archivo de Drive cada vez que alguien lo abre.
Los datos se cachean por 1 hora para no hacer demasiadas llamadas a Drive.
Cuando tú actualices el Excel en Drive, los usuarios verán los nuevos datos en máximo 1 hora.
