# Run this ON THE SERVER, elevated, if you need to remove or reinstall the service
# (e.g. after changing install_service.ps1 and re-running it).

$ServiceName = "PDFConvertApp"
$NssmExe     = "C:\nssm\nssm.exe"

Stop-Service $ServiceName -ErrorAction SilentlyContinue
& $NssmExe remove $ServiceName confirm
