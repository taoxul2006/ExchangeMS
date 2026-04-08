$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

Write-Host "Starting Exchange web app from $scriptDir" -ForegroundColor Cyan
Write-Host "Open http://127.0.0.1:3080 after startup." -ForegroundColor Yellow

& "$scriptDir\exchange-web-server.ps1"
