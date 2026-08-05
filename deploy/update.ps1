# Run this ON THE SERVER, elevated, whenever you've pushed new code and want
# the running app updated.
#
# If outbound access to github.com has been granted:
#   .\update.ps1
#   -> pulls latest code, updates any changed dependencies, restarts the service.
#
# If still offline:
#   1. Copy the updated files (or a fresh zip of the repo) into $AppDir yourself first.
#   2. If requirements.txt changed, also drop the new wheels into $AppDir\packages
#      (built locally with: pip download -r requirements.txt -d packages).
#   3. Then run:  .\update.ps1 -Offline
#      -> skips git pull, installs from the local packages folder, restarts the service.
#
# .env and the Google service-account json are gitignored, so neither path touches them.

param(
    [switch]$Offline
)

$ServiceName = "PDFConvertApp"
$AppDir      = "C:\apps\pdfconvert"
$VenvDir     = "$AppDir\.venv"

Set-Location $AppDir

Write-Host "Stopping service..."
Stop-Service $ServiceName

if (-not $Offline) {
    Write-Host "Pulling latest code..."
    git pull
    Write-Host "Installing/updating dependencies..."
    & "$VenvDir\Scripts\pip.exe" install -r requirements.txt
} else {
    Write-Host "Offline mode: skipping git pull." -ForegroundColor Yellow
    Write-Host "Make sure you already copied the updated files into $AppDir before running this." -ForegroundColor Yellow
    if (Test-Path "$AppDir\packages") {
        Write-Host "Installing dependencies from local packages folder..."
        & "$VenvDir\Scripts\pip.exe" install --no-index --find-links="$AppDir\packages" -r requirements.txt
    }
}

Write-Host "Restarting service..."
Start-Service $ServiceName
Start-Sleep -Seconds 3
Get-Service $ServiceName

Write-Host "`nCheck the logs to confirm it came up clean:"
Write-Host "Get-Content $AppDir\logs\service.err.log -Tail 30"
