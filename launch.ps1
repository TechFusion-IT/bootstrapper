$host.UI.RawUI.WindowTitle = 'TechFusion IT - Bootstrap'

# Hide the console window immediately
Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("User32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'
[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0) | Out-Null

$ErrorActionPreference = 'Stop'

$scriptUrl = 'https://raw.githubusercontent.com/TechFusion-IT/bootstrapper/main/System_Bootstrapper.ps1'

try {
    $tmp = Join-Path $env:TEMP "TFBootstrap_$(Get-Random).ps1"
    Invoke-RestMethod -Uri $scriptUrl -UseBasicParsing | Out-File -FilePath $tmp -Encoding UTF8 -Force

    Start-Process powershell.exe -ArgumentList @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-STA',
        '-WindowStyle', 'Hidden',
        '-File', $tmp
    ) -Wait

} catch {
    [System.Windows.Forms.MessageBox]::Show("Bootstrap failed:`n$_", "TechFusion IT", 0, 16)
} finally {
    if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}

[System.Environment]::Exit(0)
