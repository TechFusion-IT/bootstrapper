# ==============================================================================
#  TechFusion IT - Bootstrap Launcher
#
#  Usage (run as Administrator):
#    irm https://raw.githubusercontent.com/JOUW-ORG/JOUW-REPO/main/launch.ps1 | iex
#
#  Or with a short domain (e.g. Cloudflare redirect):
#    irm https://go.techfusion-it.com/bootstrap | iex
# ==============================================================================

$ErrorActionPreference = 'Stop'

# URL to the actual bootstrap script on GitHub
# Update this after you create your GitHub repository
$scriptUrl = 'https://raw.githubusercontent.com/TechFusion-IT/bootstrapper/main/System_Bootstrapper.ps1'

Write-Host ""
Write-Host "  TechFusion IT - System Bootstrap" -ForegroundColor Cyan
Write-Host "  Downloading..." -ForegroundColor DarkGray

try {
    # Download the script to a temp file
    # (cannot run WPF directly via iex - requires a separate STA process)
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
    # Clean up temp file
    if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}
