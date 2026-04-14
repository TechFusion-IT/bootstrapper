$ErrorActionPreference = 'Stop'

# URL to the bootstrap script on GitHub
$scriptUrl = 'https://raw.githubusercontent.com/TechFusion-IT/bootstrapper/main/System_Bootstrapper.ps1'

Write-Host ""
Write-Host "  TechFusion IT - System Bootstrap" -ForegroundColor Cyan
Write-Host "  Downloading..." -ForegroundColor DarkGray

try {
    $tmp = Join-Path $env:TEMP "TFBootstrap_$(Get-Random).ps1"
    Invoke-RestMethod -Uri $scriptUrl -UseBasicParsing | Out-File -FilePath $tmp -Encoding UTF8 -Force

    Write-Host "  Launching..." -ForegroundColor DarkGray
    Write-Host ""

    # Start in a new STA process - required for WPF GUI
    Start-Process powershell.exe -ArgumentList @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-STA',
        '-File', $tmp
    ) -Wait

} catch {
    Write-Host ""
    Write-Host "  ERROR: $_" -ForegroundColor Red
    Write-Host ""
    pause
} finally {
    if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}
