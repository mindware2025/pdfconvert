# Run this ON THE SERVER (DXBMWBOTA1), in an elevated (Run as Administrator) PowerShell.
#
# Prerequisites before running this script:
#   1. Code cloned to $AppDir (see step 2 in the deployment guide)
#   2. Python venv created and `pip install -r requirements.txt` already run inside it
#   3. .env and tool-mindware-0d87ca5562ad.json copied into $AppDir
#   4. NSSM downloaded from https://nssm.cc/download and nssm.exe placed at $NssmExe
#      (grab the win64 build, e.g. nssm-2.24\win64\nssm.exe)
#
# Edit the variables below if your paths/port differ, then run:
#   .\install_service.ps1

$ServiceName = "PDFConvertApp"
$AppDir      = "C:\apps\pdfconvert"
$VenvDir     = "$AppDir\.venv"
$NssmExe     = "C:\nssm\nssm.exe"
$Port        = 8501

if (-not (Test-Path $NssmExe)) {
    Write-Error "nssm.exe not found at $NssmExe. Download it from https://nssm.cc/download and update `$NssmExe."
    exit 1
}
if (-not (Test-Path "$VenvDir\Scripts\streamlit.exe")) {
    Write-Error "streamlit.exe not found in $VenvDir. Create the venv and run 'pip install -r requirements.txt' first."
    exit 1
}

New-Item -ItemType Directory -Force -Path "$AppDir\logs" | Out-Null

& $NssmExe install $ServiceName "$VenvDir\Scripts\streamlit.exe" "run app.py --server.port $Port --server.address 0.0.0.0 --server.headless true"
& $NssmExe set $ServiceName AppDirectory $AppDir
& $NssmExe set $ServiceName DisplayName "PDF Convert Tools (Streamlit)"
& $NssmExe set $ServiceName Start SERVICE_AUTO_START
& $NssmExe set $ServiceName AppStdout "$AppDir\logs\service.out.log"
& $NssmExe set $ServiceName AppStderr "$AppDir\logs\service.err.log"
& $NssmExe set $ServiceName AppRotateFiles 1
& $NssmExe set $ServiceName AppRotateBytes 10485760

New-NetFirewallRule -DisplayName "Streamlit - PDFConvertApp" -Direction Inbound -LocalPort $Port -Protocol TCP -Action Allow -ErrorAction SilentlyContinue | Out-Null

Start-Service $ServiceName
Start-Sleep -Seconds 3
Get-Service $ServiceName

Write-Host "`nDone. If the service shows 'Running', open http://$(hostname):$Port from another machine to confirm."
Write-Host "Logs: $AppDir\logs\service.out.log / service.err.log"
Write-Host "Manage later with: Restart-Service $ServiceName / Stop-Service $ServiceName / services.msc"
