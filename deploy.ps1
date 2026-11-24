# Uploads Code.gs, index.html and appsscript.json to the Apps Script project and
# updates the live web app deployment (same link).
#
# It always signs clasp in first: the company's Google settings make clasp's
# sign-in expire about once a day, so every deploy starts with a fresh sign-in.
# A browser page opens; approve it with the company account and the upload runs
# by itself right after.
#
# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File deploy\deploy.ps1 "what changed"
param([string]$Description = 'update')

$ErrorActionPreference = 'Stop'
$node = 'C:\Program Files\nodejs\node.exe'
$clasp = Join-Path $env:APPDATA 'npm\node_modules\@google\clasp\build\src\index.js'
$scriptId = '1IFdcozoKIiJz8wHvf-Zl6tJSXBObywgKadOqjP1asS7qtpKm5dk5cDHa'
$deploymentId = 'AKfycbyhYY_Tei0ML5xlvncBPWy6U3tLWqhYgbGeIJO-XumaGXwhe7IWibNTspZV3QdKPH0bbQ'
$src = Split-Path -Parent $PSScriptRoot
$stage = Join-Path $env:TEMP 'rpi-inventory-deploy'

if (-not (Test-Path $node)) { throw "Node.js not found at $node" }
if (-not (Test-Path $clasp)) { throw "clasp not found. Install it with: npm install -g @google/clasp" }

Write-Host '--- 1/3 signing in to Google (approve in the browser) ---'
& $node $clasp login
if ($LASTEXITCODE) { throw 'clasp sign-in failed' }

Write-Host '--- 2/3 uploading files ---'
if (Test-Path $stage) { Remove-Item -Recurse -Force $stage }
New-Item -ItemType Directory $stage | Out-Null
@{ scriptId = $scriptId; rootDir = '' } | ConvertTo-Json | Set-Content -Encoding ascii (Join-Path $stage '.clasp.json')
Copy-Item -LiteralPath (Join-Path $src 'Code.gs') (Join-Path $stage 'Code.js')
Copy-Item -LiteralPath (Join-Path $src 'index.html') (Join-Path $stage 'index.html')
Copy-Item -LiteralPath (Join-Path $src 'appsscript.json') (Join-Path $stage 'appsscript.json')
Push-Location $stage
try {
  & $node $clasp push --force
  if ($LASTEXITCODE) { throw 'upload failed' }

  Write-Host '--- 3/3 updating the web app (same link) ---'
  & $node $clasp create-deployment --deploymentId $deploymentId --description $Description
  if ($LASTEXITCODE) { throw 'deployment update failed' }
  & $node $clasp list-deployments
} finally {
  Pop-Location
}
Write-Host 'DEPLOY-DONE'
