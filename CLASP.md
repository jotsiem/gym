# Servidor en Git con clasp

## Configuración (una vez)
1. Descomprime el zip en cualquier carpeta.
2. Abre PowerShell en esa carpeta y ejecuta:
   `powershell -ExecutionPolicy Bypass -File .\setup.ps1`
3. El script instala clasp, te pide activar la API de Apps Script, hace login de Google, clona jotsiem/gym en `~/gym`,
   copia los archivos, te pide el ID del script y hace la primera publicación + push.
   - Si Git pide login de GitHub, entra con la cuenta **jotsiem**.

## Cada cambio
`cd ~/gym` y `.\deploy.ps1 "descripción"` → sube el servidor, publica en la misma URL /exec y hace commit + push.
(Si PowerShell bloquea scripts: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` una vez.)

El repo es público: aquí no hay secretos (la clave vive en las propiedades del script, no en el código).
