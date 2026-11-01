# Publica una version nueva del servidor (Apps Script) y sube los cambios a GitHub.
# Uso:  .\deploy.ps1 "que has cambiado"
# La URL /exec no cambia: se actualiza siempre la misma implementacion.
param([string]$msg = "actualizacion")
$ErrorActionPreference = "Stop"
$deployId = "AKfycbze03FedycMRlBIguasaVBrlvlqKH3i0i7XXGghKRTSdRdUFsPssEHQ9wF9B_hk7GYq0g"
Push-Location $PSScriptRoot
try {
  clasp push --force
  if ($LASTEXITCODE -ne 0) { throw "clasp push ha fallado" }
  clasp deploy -i $deployId -d $msg
  if ($LASTEXITCODE -ne 0) { throw "clasp deploy ha fallado" }
  git add -A
  git commit -m $msg
  git push
  Write-Host "Listo: servidor publicado en la misma URL y GitHub actualizado." -ForegroundColor Green
} finally { Pop-Location }
