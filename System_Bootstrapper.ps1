# ==============================================================================
#  TechFusion IT - System Bootstrap Tool
#  Version: 1.0.0
#  Author:  TechFusion IT (https://www.techfusion-it.com)
#
#  Usage:
#    Run as Administrator:
#      Set-ExecutionPolicy -Scope Process Bypass
#      .\TechFusion-Bootstrap.ps1
#
#  Rebranding: edit only $branding and $theme below.
# ==============================================================================

#Requires -Version 5.1
Add-Type -AssemblyName PresentationFramework

# ==============================================================================
#  HIDE CONSOLE WINDOW
#  Hides the PowerShell console so only the GUI is visible.
#  The process exits cleanly when the GUI window is closed.
# ==============================================================================

Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("User32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'
[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0) | Out-Null

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ==============================================================================
#  BRANDING
# ==============================================================================

$branding = @{
    CompanyName  = "TechFusion IT"
    WindowTitle  = "TechFusion IT - System Bootstrap"
    DefaultAdmin = "TechFusion"
    Version      = "v1.0.0"
}

# ==============================================================================
#  THEME
# ==============================================================================

$theme = @{
    WindowBg          = "#0b0f1a"
    HeaderBg          = "#111827"
    CardBg            = "#111827"
    Accent            = "#00b4d8"
    AccentHover       = "#0096b5"
    NavBtn            = "#1a2744"
    NavBtnHover       = "#223057"
    GhostBorder       = "#1e2d4a"
    GhostBorderHover  = "#00b4d8"
    InputBg           = "#0d1526"
    InputBorder       = "#1e2d4a"
    PwdBoxBg          = "#071020"
    DotIdle           = "#1e2d4a"
    TextPrimary       = "#e8f4fd"
    TextSecondary     = "#7a9bb5"
    TextMuted         = "#3a5068"
    TextLabel         = "#7a9bb5"
    CategoryColor     = "#00b4d8"
    SeparatorColor    = "#1a2744"
    FontFamily        = "Segoe UI"
}

# ==============================================================================
#  APPLICATIONS
#  Find winget IDs with: winget search <name>
# ==============================================================================

$apps = @(
	@{ Category = "Browsers";		Id = "XPFFTQ037JWMHS";				Name = "Microsoft Edge" }  
    @{ Category = "Browsers";		Id = "Google.Chrome";				Name = "Google Chrome" }
    @{ Category = "Browsers";		Id = "Mozilla.Firefox";				Name = "Mozilla Firefox" }
	@{ Category = "Browsers";		Id = "XPDBZ4MPRKNN30";				Name = "Opera GX" }
    @{ Category = "Productivity";	Id = "Microsoft.Office";			Name = "Microsoft 365 / Office" }
	@{ Category = "Productivity";	Id = "Microsoft.OneDrive";			Name = "Microsoft OneDrive" }
	@{ Category = "Productivity";	Id = "Adobe.Acrobat.Reader.64-bit";	Name = "Adobe Acrobat Reader" }
    @{ Category = "Communication";	Id = "Microsoft.Teams";				Name = "Microsoft Teams" }
	@{ Category = "Communication";	Id = "9NKSQGP7F2NH";				Name = "WhatsApp" }
	@{ Category = "Communication";	Id = "9NBDXK71NK08";				Name = "WhatsApp Beta" }
	@{ Category = "AI";				Id = "9NT1R1C2HH7J";				Name = "ChatGPT" }
	@{ Category = "AI";				Id = "Anthropic.Claude";			Name = "Claude" }
	@{ Category = "AI";				Id = "Anthropic.ClaudeCode";		Name = "Claude Code" }
	@{ Category = "AI";				Id = "XP9CXNGPPJ97XX";				Name = "Microsoft Copilot" }
	@{ Category = "AI";				Id = "XP8JNQFBQH6PVF";				Name = "Perplexity" }
	@{ Category = "Media";			Id = "9WZDNCRFJ3TJ";				Name = "Netflix" }
	@{ Category = "Media";			Id = "9NXQXXLFST89";				Name = "Disney+" }
	@{ Category = "Media";			Id = "9P6RC76MSMMJ";				Name = "Prime Video" }
	@{ Category = "Media";			Id = "9NCBCSZSJRSB";				Name = "Spotify" }
    @{ Category = "Media";			Id = "VideoLAN.VLC";				Name = "VLC Media Player" }
	@{ Category = "Utilities";		Id = "7zip.7zip";					Name = "7-Zip" }
	@{ Category = "Utilities";		Id = "RARLab.WinRAR";				Name = "Winrar" }
	@{ Category = "Utilities";		Id = "9N09F8V8FS02";				Name = "DisplayLink Manager" }
	@{ Category = "Utilities";		Id = "9P8K5G2MWW6Z";				Name = "Intel Graphics Software" }
    @{ Category = "Utilities";		Id = "Notepad++.Notepad++";			Name = "Notepad++" }
	@{ Category = "Utilities";		Id = "dotPDN.PaintDotNet";			Name = "Paint.Net" }
	@{ Category = "Utilities";		Id = "9NBLGGH4Z1JC";				Name = "Ookla Speedtest" }
	@{ Category = "Security";		Id = "9P6PMZTM93LR";				Name = "Microsoft Defender" }
    @{ Category = "Security";		Id = "XPFM2WPGW7NLVB";				Name = "Eset Home Essential" }
	@{ Category = "Security";		Id = "XP9K0R82SCX2PM";				Name = "Eset Home Premium" }
	@{ Category = "Security";		Id = "XP9KF40VGV9PWM";				Name = "Eset Nod32 Antivirus" }
)

# ==============================================================================
#  TWEAKS
#  Add a matching case in Invoke-Tweak for every entry added here.
# ==============================================================================

$tweaks = @(
    @{ Id = "DarkMode";				Name = "Enable Dark Mode";					Description = "Switches Windows to dark theme (apps and system)" }
    @{ Id = "ShowFileExt";			Name = "Show file extensions";				Description = "Shows .exe .docx etc. in File Explorer" }
    @{ Id = "ShowHidden";			Name = "Show hidden files and folders";		Description = "Reveals hidden items in File Explorer" }
    @{ Id = "DisableTelemetry";		Name = "Disable telemetry";					Description = "Reduces diagnostic data sent to Microsoft" }
    @{ Id = "ClassicContext";		Name = "Restore classic right-click menu";	Description = "Brings back the full context menu in Windows 11" }
    @{ Id = "DisableBing";			Name = "Disable Bing in Start Menu";		Description = "Search Start Menu locally only" }
	@{ Id = "DisableNewOutlook";	Name = "Disable new Outlook";				Description = "Hides the toggle and blocks auto-migration to new Outlook" }
    @{ Id = "NumLock";				Name = "Enable NumLock on startup";			Description = "NumLock is always active after boot" }
    @{ Id = "ClipboardHistory";		Name = "Enable Clipboard History";			Description = "Enables Win+V multi-item clipboard history" }
)

# ==============================================================================
#  TWEAK IMPLEMENTATIONS
# ==============================================================================

function Invoke-Tweak {
    param([string]$Id)
    switch ($Id) {
        "DarkMode" {
            $p = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            Set-ItemProperty -Path $p -Name AppsUseLightTheme    -Value 0 -Type DWord -Force
            Set-ItemProperty -Path $p -Name SystemUsesLightTheme -Value 0 -Type DWord -Force
        }
        "ShowFileExt" {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name HideFileExt -Value 0 -Type DWord -Force
        }
        "ShowHidden" {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name Hidden -Value 1 -Type DWord -Force
        }
        "DisableTelemetry" {
            $p = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
            if (-not (Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
            Set-ItemProperty -Path $p -Name AllowTelemetry -Value 0 -Type DWord -Force
        }
        "ClassicContext" {
            $p = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
            if (-not (Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
            Set-ItemProperty -Path $p -Name "(Default)" -Value "" -Force
        }
        "DisableBing" {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name BingSearchEnabled -Value 0 -Type DWord -Force
        }
        "NumLock" {
            Set-ItemProperty -Path "HKCU:\Control Panel\Keyboard" -Name InitialKeyboardIndicators -Value 2 -Force
        }
        "ClipboardHistory" {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Clipboard" -Name EnableClipboardHistory -Value 1 -Type DWord -Force
        }
		"DisableNewOutlook" {
            # Hide the new Outlook toggle in classic Outlook
            $p1 = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General"
            if (-not (Test-Path $p1)) { New-Item -Path $p1 -Force | Out-Null }
            Set-ItemProperty -Path $p1 -Name HideNewOutlookToggle -Value 1 -Type DWord -Force

            # Block automatic migration to new Outlook
            $p2 = "HKCU:\Software\Policies\Microsoft\office\16.0\outlook\preferences"
            if (-not (Test-Path $p2)) { New-Item -Path $p2 -Force | Out-Null }
            Set-ItemProperty -Path $p2 -Name NewOutlookMigrationUserSetting    -Value 0 -Type DWord -Force
            Set-ItemProperty -Path $p2 -Name NewOutlookAutomaticSetupUserSetting -Value 0 -Type DWord -Force

            # Also set via policies path for the toggle
            $p3 = "HKCU:\Software\Policies\Microsoft\office\16.0\outlook\options\general"
            if (-not (Test-Path $p3)) { New-Item -Path $p3 -Force | Out-Null }
            Set-ItemProperty -Path $p3 -Name HideNewOutlookToggle -Value 1 -Type DWord -Force
        }
    }
}

# ==============================================================================
#  PASSWORD GENERATOR
#  Exactly 12 characters. Guaranteed: 1 uppercase, 1 lowercase, 1 digit,
#  1 special character. Remaining 8 positions filled randomly. Fisher-Yates
#  shuffled so the guaranteed chars are not always in positions 0-3.
# ==============================================================================

function New-RandomPassword {
    $upper   = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $lower   = "abcdefghijklmnopqrstuvwxyz"
    $digits  = "0123456789"
    $special = "!@#$%^&*()-_=+"
    $all     = $upper + $lower + $digits + $special
    $rng     = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $buf     = [byte[]]::new(1)
    function Get-RC { param([string]$cs); $rng.GetBytes($buf); $cs[$buf[0] % $cs.Length] }

    # 4 guaranteed + 8 random = exactly 12
    $pwd = [System.Collections.Generic.List[char]]::new()
    $pwd.Add((Get-RC $upper))
    $pwd.Add((Get-RC $lower))
    $pwd.Add((Get-RC $digits))
    $pwd.Add((Get-RC $special))
    for ($i = 4; $i -lt 12; $i++) { $pwd.Add((Get-RC $all)) }

    # Fisher-Yates shuffle
    for ($i = 11; $i -gt 0; $i--) {
        $rng.GetBytes($buf)
        $j   = $buf[0] % ($i + 1)
        $tmp = $pwd[$i]; $pwd[$i] = $pwd[$j]; $pwd[$j] = $tmp
    }
    return ($pwd -join "")
}

# ==============================================================================
#  XAML TEMPLATE
# ==============================================================================

$xamlTemplate = '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="{{WindowTitle}}" Width="760" Height="680" MinWidth="660" MinHeight="560" WindowStartupLocation="CenterScreen" Background="{{WindowBg}}" FontFamily="{{FontFamily}}"><Window.Resources><Style TargetType="CheckBox"><Setter Property="Foreground" Value="{{TextPrimary}}"/><Setter Property="Margin" Value="0,3,0,3"/><Setter Property="FontSize" Value="13"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style><Style TargetType="TextBox" x:Key="InputBox"><Setter Property="Background" Value="{{InputBg}}"/><Setter Property="Foreground" Value="{{TextPrimary}}"/><Setter Property="CaretBrush" Value="{{TextPrimary}}"/><Setter Property="BorderBrush" Value="{{InputBorder}}"/><Setter Property="BorderThickness" Value="1"/><Setter Property="FontSize" Value="13"/><Setter Property="Padding" Value="8,7"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style><Style TargetType="PasswordBox" x:Key="PwdBox"><Setter Property="Background" Value="{{InputBg}}"/><Setter Property="Foreground" Value="{{TextPrimary}}"/><Setter Property="CaretBrush" Value="{{TextPrimary}}"/><Setter Property="BorderBrush" Value="{{InputBorder}}"/><Setter Property="BorderThickness" Value="1"/><Setter Property="FontSize" Value="13"/><Setter Property="Padding" Value="8,7"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style><Style TargetType="Button" x:Key="NavBtn"><Setter Property="Background" Value="{{NavBtn}}"/><Setter Property="Foreground" Value="{{TextPrimary}}"/><Setter Property="FontSize" Value="13"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Padding" Value="18,8"/><Setter Property="BorderThickness" Value="0"/><Setter Property="Cursor" Value="Hand"/><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{{NavBtnHover}}"/></Trigger><Trigger Property="IsEnabled" Value="False"><Setter Property="Background" Value="#0d1526"/><Setter Property="Foreground" Value="#3a5068"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style><Style TargetType="Button" x:Key="AccentBtn"><Setter Property="Background" Value="{{Accent}}"/><Setter Property="Foreground" Value="#0b0f1a"/><Setter Property="FontSize" Value="13"/><Setter Property="FontWeight" Value="Bold"/><Setter Property="Padding" Value="20,9"/><Setter Property="BorderThickness" Value="0"/><Setter Property="Cursor" Value="Hand"/><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{{AccentHover}}"/></Trigger><Trigger Property="IsEnabled" Value="False"><Setter Property="Background" Value="#0d1526"/><Setter Property="Foreground" Value="#3a5068"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style><Style TargetType="Button" x:Key="GhostBtn"><Setter Property="Background" Value="Transparent"/><Setter Property="Foreground" Value="{{TextSecondary}}"/><Setter Property="FontSize" Value="13"/><Setter Property="Padding" Value="18,8"/><Setter Property="BorderThickness" Value="1"/><Setter Property="BorderBrush" Value="{{GhostBorder}}"/><Setter Property="Cursor" Value="Hand"/><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="{{TextPrimary}}"/><Setter Property="BorderBrush" Value="{{GhostBorderHover}}"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style></Window.Resources><Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions><Border Grid.Row="0" Background="{{HeaderBg}}" Padding="24,16"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><StackPanel Grid.Column="0" VerticalAlignment="Center"><TextBlock x:Name="lblStepTitle" Text="Step 1 of 3" FontSize="18" FontWeight="Bold" Foreground="{{TextPrimary}}"/><TextBlock x:Name="lblStepSub" Text="" Foreground="{{TextSecondary}}" FontSize="12" Margin="0,3,0,0"/></StackPanel><StackPanel Grid.Column="1" HorizontalAlignment="Right"><TextBlock Text="{{CompanyName}}" Foreground="{{Accent}}" FontSize="14" FontWeight="Bold" HorizontalAlignment="Right"/><TextBlock x:Name="lblVersion" Text="{{Version}}" Foreground="{{TextMuted}}" FontSize="10" HorizontalAlignment="Right" Margin="0,2,0,0"/></StackPanel></Grid></Border><Border Grid.Row="1" Background="{{HeaderBg}}" Padding="24,0,24,14"><StackPanel Orientation="Horizontal"><Ellipse x:Name="dot1" Width="10" Height="10" Fill="{{Accent}}" Margin="0,0,6,0"/><TextBlock x:Name="lbl1" Text="Admin" Foreground="{{Accent}}" FontSize="11" VerticalAlignment="Center" Margin="0,0,12,0"/><Rectangle Width="24" Height="1" Fill="{{DotIdle}}" VerticalAlignment="Center" Margin="0,0,12,0"/><Ellipse x:Name="dot2" Width="10" Height="10" Fill="{{DotIdle}}" Margin="0,0,6,0"/><TextBlock x:Name="lbl2" Text="Apps" Foreground="{{TextMuted}}" FontSize="11" VerticalAlignment="Center" Margin="0,0,12,0"/><Rectangle Width="24" Height="1" Fill="{{DotIdle}}" VerticalAlignment="Center" Margin="0,0,12,0"/><Ellipse x:Name="dot3" Width="10" Height="10" Fill="{{DotIdle}}" Margin="0,0,6,0"/><TextBlock x:Name="lbl3" Text="Tweaks" Foreground="{{TextMuted}}" FontSize="11" VerticalAlignment="Center"/></StackPanel></Border><Grid Grid.Row="2" Margin="24,18,24,8"><StackPanel x:Name="frameStep1" Visibility="Visible"><TextBlock Text="LOCAL ADMINISTRATOR ACCOUNT" Foreground="{{TextSecondary}}" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,10"/><CheckBox x:Name="chkCreateAdmin" Content="Create or update a local administrator account on this machine" FontSize="13" FontWeight="SemiBold" Foreground="{{TextPrimary}}" Margin="0,0,0,18"/><StackPanel x:Name="pnlAdminFields" Visibility="Collapsed"><Border Background="{{CardBg}}" CornerRadius="6" Padding="18,14"><StackPanel><TextBlock Text="USERNAME" Foreground="{{TextMuted}}" FontSize="10" FontWeight="Bold" Margin="0,0,0,5"/><TextBox x:Name="txtUsername" Style="{StaticResource InputBox}" Margin="0,0,0,16"/><TextBlock Text="PASSWORD" Foreground="{{TextMuted}}" FontSize="10" FontWeight="Bold" Margin="0,0,0,5"/><Grid Margin="0,0,0,12"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><PasswordBox x:Name="pwdPassword" Style="{StaticResource PwdBox}" Grid.Column="0" Margin="0,0,8,0"/><Button x:Name="btnGenerate" Grid.Column="1" Content="Generate" Style="{StaticResource NavBtn}"/></Grid><Border x:Name="pnlGeneratedPwd" Visibility="Collapsed" Background="{{PwdBoxBg}}" CornerRadius="5" Padding="14,12" Margin="0,0,0,12"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><StackPanel Grid.Column="0" VerticalAlignment="Center"><TextBlock Text="GENERATED PASSWORD - COPY BEFORE CONTINUING" Foreground="{{Accent}}" FontSize="10" FontWeight="Bold" Margin="0,0,0,6"/><TextBlock x:Name="lblGeneratedPwd" Text="" Foreground="{{TextPrimary}}" FontFamily="Consolas" FontSize="18" FontWeight="Bold"/></StackPanel><Button x:Name="btnCopyPwd" Grid.Column="1" Content="Copy" Style="{StaticResource AccentBtn}" VerticalAlignment="Center" Margin="12,0,0,0"/></Grid></Border><TextBlock Text="Tip: Click Generate for a secure random password, or type your own above." Foreground="{{TextMuted}}" FontSize="11" TextWrapping="Wrap"/></StackPanel></Border></StackPanel></StackPanel><Grid x:Name="frameStep2" Visibility="Collapsed"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions><Grid Grid.Row="0" Margin="0,0,0,10"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" Text="APPLICATIONS" Foreground="{{TextSecondary}}" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/><Border Grid.Column="1" x:Name="toggleAppsBorder" Background="{{NavBtn}}" CornerRadius="4" Padding="0"><StackPanel Orientation="Horizontal"><Button x:Name="btnModeInstall" Content="Install / Update" FontSize="11" FontWeight="SemiBold" Padding="10,5" BorderThickness="0" Cursor="Hand" Background="{{Accent}}" Foreground="#0b0f1a"><Button.Template><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" CornerRadius="4,0,0,4" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border></ControlTemplate></Button.Template></Button><Button x:Name="btnModeUninstall" Content="Uninstall" FontSize="11" FontWeight="SemiBold" Padding="10,5" BorderThickness="0" Cursor="Hand" Background="Transparent" Foreground="{{TextSecondary}}"><Button.Template><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" CornerRadius="0,4,4,0" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border></ControlTemplate></Button.Template></Button></StackPanel></Border></Grid><Border Grid.Row="1" Background="{{CardBg}}" CornerRadius="6" Padding="12"><Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions><StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8"><Button x:Name="btnSelectAllApps" Content="Select all" Style="{StaticResource NavBtn}" Margin="0,0,8,0"/><Button x:Name="btnDeselectAllApps" Content="Deselect all" Style="{StaticResource NavBtn}"/></StackPanel><ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" MaxHeight="380"><StackPanel x:Name="pnlApps"/></ScrollViewer></Grid></Border></Grid><Grid x:Name="frameStep3" Visibility="Collapsed"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions><Grid Grid.Row="0" Margin="0,0,0,10"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Grid.Column="0" Text="SYSTEM TWEAKS" Foreground="{{TextSecondary}}" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/><Border Grid.Column="1" x:Name="toggleTweaksBorder" Background="{{NavBtn}}" CornerRadius="4" Padding="0"><StackPanel Orientation="Horizontal"><Button x:Name="btnModeApply" Content="Apply" FontSize="11" FontWeight="SemiBold" Padding="10,5" BorderThickness="0" Cursor="Hand" Background="{{Accent}}" Foreground="#0b0f1a"><Button.Template><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" CornerRadius="4,0,0,4" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border></ControlTemplate></Button.Template></Button><Button x:Name="btnModeUndo" Content="Undo / Revert" FontSize="11" FontWeight="SemiBold" Padding="10,5" BorderThickness="0" Cursor="Hand" Background="Transparent" Foreground="{{TextSecondary}}"><Button.Template><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" CornerRadius="0,4,4,0" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border></ControlTemplate></Button.Template></Button></StackPanel></Border></Grid><Border Grid.Row="1" Background="{{CardBg}}" CornerRadius="6" Padding="12"><ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="380"><StackPanel x:Name="pnlTweaks"/></ScrollViewer></Border></Grid><Grid x:Name="frameStep4" Visibility="Collapsed"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions><TextBlock Grid.Row="0" Text="OUTPUT LOG" Foreground="{{TextSecondary}}" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,8"/><Border Grid.Row="1" Background="{{CardBg}}" CornerRadius="6" Padding="14"><ScrollViewer x:Name="logScroller" VerticalScrollBarVisibility="Auto"><TextBlock x:Name="txtLog" Foreground="{{TextSecondary}}" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap"/></ScrollViewer></Border></Grid></Grid><Border Grid.Row="3" Background="{{HeaderBg}}" Padding="24,12"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock x:Name="lblStatus" Text="Step 1 of 3" Foreground="{{TextMuted}}" VerticalAlignment="Center" FontSize="12"/><StackPanel Grid.Column="1" Orientation="Horizontal"><Button x:Name="btnSkip" Content="Skip this step" Style="{StaticResource GhostBtn}" Margin="0,0,8,0"/><Button x:Name="btnNext" Content="Next step" Style="{StaticResource AccentBtn}"/></StackPanel></Grid></Border></Grid></Window>'

# ==============================================================================
#  INJECT THEME + BRANDING
# ==============================================================================

$xamlString = $xamlTemplate
foreach ($key in $theme.Keys)    { $xamlString = $xamlString.Replace("{{$key}}", $theme[$key]) }
foreach ($key in $branding.Keys) { $xamlString = $xamlString.Replace("{{$key}}", $branding[$key]) }

[xml]$xaml = $xamlString
$reader    = [System.Xml.XmlNodeReader]::new($xaml)
$window    = [Windows.Markup.XamlReader]::Load($reader)

# ==============================================================================
#  RESOLVE CONTROLS
# ==============================================================================

function Get-C { param([string]$n); $window.FindName($n) }

$lblStepTitle       = Get-C "lblStepTitle"
$lblStepSub         = Get-C "lblStepSub"
$lblStatus          = Get-C "lblStatus"
$dot1 = Get-C "dot1"; $dot2 = Get-C "dot2"; $dot3 = Get-C "dot3"
$lbl1 = Get-C "lbl1"; $lbl2 = Get-C "lbl2"; $lbl3 = Get-C "lbl3"
$frameStep1         = Get-C "frameStep1"
$frameStep2         = Get-C "frameStep2"
$frameStep3         = Get-C "frameStep3"
$frameStep4         = Get-C "frameStep4"
$chkCreateAdmin     = Get-C "chkCreateAdmin"
$pnlAdminFields     = Get-C "pnlAdminFields"
$txtUsername        = Get-C "txtUsername"
$pwdPassword        = Get-C "pwdPassword"
$btnGenerate        = Get-C "btnGenerate"
$pnlGeneratedPwd    = Get-C "pnlGeneratedPwd"
$lblGeneratedPwd    = Get-C "lblGeneratedPwd"
$btnCopyPwd         = Get-C "btnCopyPwd"
$pnlApps            = Get-C "pnlApps"
$btnSelectAllApps   = Get-C "btnSelectAllApps"
$btnDeselectAllApps = Get-C "btnDeselectAllApps"
$btnModeInstall     = Get-C "btnModeInstall"
$btnModeUninstall   = Get-C "btnModeUninstall"
$btnModeApply       = Get-C "btnModeApply"
$btnModeUndo        = Get-C "btnModeUndo"
$pnlTweaks          = Get-C "pnlTweaks"
$txtLog             = Get-C "txtLog"
$logScroller        = Get-C "logScroller"
$btnSkip            = Get-C "btnSkip"
$btnNext            = Get-C "btnNext"

# ==============================================================================
#  HELPERS
# ==============================================================================

function New-Brush { param([string]$h); [System.Windows.Media.BrushConverter]::new().ConvertFromString($h) }
$Collapsed = [System.Windows.Visibility]::Collapsed
$Visible   = [System.Windows.Visibility]::Visible

# ==============================================================================
#  MODE TOGGLES
#  appMode:   'install' | 'uninstall'
#  tweakMode: 'apply'   | 'undo'
# ==============================================================================

$script:appMode   = 'install'
$script:tweakMode = 'apply'

function Set-AppMode {
    param([string]$Mode)
    $script:appMode = $Mode
    if ($Mode -eq 'install') {
        $btnModeInstall.Background   = New-Brush $theme.Accent
        $btnModeInstall.Foreground   = New-Brush '#0b0f1a'
        $btnModeUninstall.Background = New-Brush 'Transparent'
        $btnModeUninstall.Foreground = New-Brush $theme.TextSecondary
        $lblStepSub.Text = 'Select the applications to install or update via winget.'
    } else {
        $btnModeUninstall.Background = New-Brush '#c0392b'
        $btnModeUninstall.Foreground = New-Brush 'White'
        $btnModeInstall.Background   = New-Brush 'Transparent'
        $btnModeInstall.Foreground   = New-Brush $theme.TextSecondary
        $lblStepSub.Text = 'UNINSTALL MODE  -  selected apps will be removed from this system.'
    }
}

function Set-TweakMode {
    param([string]$Mode)
    $script:tweakMode = $Mode
    if ($Mode -eq 'apply') {
        $btnModeApply.Background = New-Brush $theme.Accent
        $btnModeApply.Foreground = New-Brush '#0b0f1a'
        $btnModeUndo.Background  = New-Brush 'Transparent'
        $btnModeUndo.Foreground  = New-Brush $theme.TextSecondary
        $lblStepSub.Text = 'Select system settings and registry tweaks to apply.'
    } else {
        $btnModeUndo.Background  = New-Brush '#c0392b'
        $btnModeUndo.Foreground  = New-Brush 'White'
        $btnModeApply.Background = New-Brush 'Transparent'
        $btnModeApply.Foreground = New-Brush $theme.TextSecondary
        $lblStepSub.Text = 'UNDO MODE  -  selected tweaks will be reverted to Windows defaults.'
    }
}

$btnModeInstall.Add_Click({   Set-AppMode 'install' })
$btnModeUninstall.Add_Click({ Set-AppMode 'uninstall' })
$btnModeApply.Add_Click({     Set-TweakMode 'apply' })
$btnModeUndo.Add_Click({      Set-TweakMode 'undo' })

# ==============================================================================
#  STEP INDICATOR
# ==============================================================================

$script:currentStep = 1

$stepTitles = @{
    1 = "Step 1 of 3  -  Local Administrator Account"
    2 = "Step 2 of 3  -  Install Applications"
    3 = "Step 3 of 3  -  System Tweaks"
    4 = "Ready to Execute"
}
$stepSubs = @{
    1 = "Create or update a local admin account for this machine."
    2 = "Select the applications to install or update via winget."
    3 = "Select system settings and registry tweaks to apply."
    4 = "All selected actions will now be executed."
}

function Update-Dots {
    $s  = $script:currentStep
    $cA = New-Brush $theme.Accent
    $cD = New-Brush "#22863a"
    $cI = New-Brush $theme.DotIdle
    $tA = New-Brush $theme.Accent
    $tD = New-Brush "#22863a"
    $tI = New-Brush $theme.TextMuted
    $dot1.Fill = if ($s -gt 1) { $cD } elseif ($s -eq 1) { $cA } else { $cI }
    $dot2.Fill = if ($s -gt 2) { $cD } elseif ($s -eq 2) { $cA } else { $cI }
    $dot3.Fill = if ($s -gt 3) { $cD } elseif ($s -eq 3) { $cA } else { $cI }
    $lbl1.Foreground = if ($s -gt 1) { $tD } elseif ($s -eq 1) { $tA } else { $tI }
    $lbl2.Foreground = if ($s -gt 2) { $tD } elseif ($s -eq 2) { $tA } else { $tI }
    $lbl3.Foreground = if ($s -gt 3) { $tD } elseif ($s -eq 3) { $tA } else { $tI }
}

function Show-Step {
    param([int]$Step)
    $script:currentStep = $Step
    $frameStep1.Visibility = $Collapsed
    $frameStep2.Visibility = $Collapsed
    $frameStep3.Visibility = $Collapsed
    $frameStep4.Visibility = $Collapsed
    switch ($Step) {
        1 { $frameStep1.Visibility = $Visible }
        2 { $frameStep2.Visibility = $Visible }
        3 { $frameStep3.Visibility = $Visible }
        4 { $frameStep4.Visibility = $Visible; $btnSkip.Visibility = $Collapsed; $btnNext.Content = "Execute" }
    }
    $lblStepTitle.Text = $stepTitles[$Step]
    $lblStepSub.Text   = $stepSubs[$Step]
    $lblStatus.Text    = if ($Step -le 3) { "Step $Step of 3" } else { "Ready to execute" }
    Update-Dots
}

# ==============================================================================
#  STEP 1 - ADMIN
# ==============================================================================

$txtUsername.Text = $branding.DefaultAdmin

$chkCreateAdmin.Add_Checked({   $pnlAdminFields.Visibility = $Visible })
$chkCreateAdmin.Add_Unchecked({ $pnlAdminFields.Visibility = $Collapsed })

$btnGenerate.Add_Click({
    $pwd = New-RandomPassword
    $lblGeneratedPwd.Text       = $pwd
    $pnlGeneratedPwd.Visibility = $Visible
    $pwdPassword.Password       = $pwd
})

$btnCopyPwd.Add_Click({
    [System.Windows.Clipboard]::SetText($lblGeneratedPwd.Text)
    $btnCopyPwd.Content = "Copied!"
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromSeconds(2)
    $timer.Add_Tick({
        $timer.Stop()
        if ($frameStep1.Visibility -eq $Visible -and $null -ne $btnCopyPwd) {
            $btnCopyPwd.Content = "Copy"
        }
    })
    $timer.Start()
})

$pwdPassword.Add_PasswordChanged({
    if ($pwdPassword.Password -ne $lblGeneratedPwd.Text) {
        $pnlGeneratedPwd.Visibility = $Collapsed
    }
})

# ==============================================================================
#  STEP 2 - APPS
# ==============================================================================

$appCheckboxes   = @{}
$currentCategory = ""

foreach ($app in $apps) {
    if ($app.Category -ne $currentCategory) {
        $currentCategory = $app.Category
        $hdr = [System.Windows.Controls.TextBlock]::new()
        $hdr.Text       = $currentCategory.ToUpper()
        $hdr.Foreground = New-Brush $theme.CategoryColor
        $hdr.FontSize   = 10
        $hdr.FontWeight = [System.Windows.FontWeights]::Bold
        $hdr.Margin     = [System.Windows.Thickness]::new(4,10,0,3)
        $pnlApps.Children.Add($hdr) | Out-Null
        $sep = [System.Windows.Controls.Separator]::new()
        $sep.Background = New-Brush $theme.SeparatorColor
        $sep.Margin     = [System.Windows.Thickness]::new(0,0,0,4)
        $pnlApps.Children.Add($sep) | Out-Null
    }
    $cb = [System.Windows.Controls.CheckBox]::new()
    $cb.Content    = $app.Name
    $cb.Tag        = $app.Id
    $cb.Foreground = New-Brush $theme.TextPrimary
    $cb.Margin     = [System.Windows.Thickness]::new(4,3,0,3)
    $cb.FontSize   = 13
    $pnlApps.Children.Add($cb) | Out-Null
    $appCheckboxes[$app.Id] = $cb
}

$btnSelectAllApps.Add_Click({   foreach ($cb in $appCheckboxes.Values) { $cb.IsChecked = $true } })
$btnDeselectAllApps.Add_Click({ foreach ($cb in $appCheckboxes.Values) { $cb.IsChecked = $false } })

# ==============================================================================
#  STEP 3 - TWEAKS
# ==============================================================================

$tweakCheckboxes = @{}

foreach ($tweak in $tweaks) {
    $panel = [System.Windows.Controls.StackPanel]::new()
    $panel.Margin = [System.Windows.Thickness]::new(4,4,0,4)
    $cb = [System.Windows.Controls.CheckBox]::new()
    $cb.Content    = $tweak.Name
    $cb.Tag        = $tweak.Id
    $cb.Foreground = New-Brush $theme.TextPrimary
    $cb.FontSize   = 13
    $desc = [System.Windows.Controls.TextBlock]::new()
    $desc.Text       = "    " + $tweak.Description
    $desc.Foreground = New-Brush $theme.TextMuted
    $desc.FontSize   = 11
    $desc.Margin     = [System.Windows.Thickness]::new(0,2,0,0)
    $panel.Children.Add($cb)   | Out-Null
    $panel.Children.Add($desc) | Out-Null
    $pnlTweaks.Children.Add($panel) | Out-Null
    $tweakCheckboxes[$tweak.Id] = $cb
}

# ==============================================================================
#  NAVIGATION
# ==============================================================================

$btnSkip.Add_Click({ Show-Step ($script:currentStep + 1) })

$btnNext.Add_Click({
    if ($script:currentStep -lt 4) {
        if ($script:currentStep -eq 1 -and $chkCreateAdmin.IsChecked -eq $true) {
            if ([string]::IsNullOrWhiteSpace($txtUsername.Text)) {
                [System.Windows.MessageBox]::Show("Please enter a username.", "Missing input", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }
            if ([string]::IsNullOrWhiteSpace($pwdPassword.Password)) {
                [System.Windows.MessageBox]::Show("Please generate or type a password.", "Missing input", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }
        }
        Show-Step ($script:currentStep + 1)
    } else {
        Start-Bootstrap
    }
})

# ==============================================================================
#  UI LOG
# ==============================================================================

function Write-UILog {
    param([string]$Msg, [string]$Colour = "#7a9bb5")
    $ts  = Get-Date -Format "HH:mm:ss"
    $run = [System.Windows.Documents.Run]::new("[$ts] $Msg`n")
    $run.Foreground = New-Brush $Colour
    $txtLog.Inlines.Add($run)
    $logScroller.ScrollToBottom()
}

# ==============================================================================
#  EXECUTE
#  Runs in a separate STA runspace so the UI stays responsive.
#  IMPORTANT: Dispatcher.Invoke (blocking) is used inside the worker so every
#  log line finishes writing before the next registry call starts. This also
#  prevents the runspace from exiting before the final UI update completes.
# ==============================================================================

function Start-Bootstrap {
    $btnNext.IsEnabled  = $false
    $btnSkip.Visibility = $Collapsed

    $doAdmin  = $chkCreateAdmin.IsChecked -eq $true
    $uname    = $txtUsername.Text
    $upwd     = $pwdPassword.Password
    # Steps are executed when reached via Next (skip = bypass entirely)
    $doApps   = $true
    $doTweaks = $true

    $appsToInstall = @($appCheckboxes.GetEnumerator()   | Where-Object { $_.Value.IsChecked -eq $true } | ForEach-Object { $_.Key })
    $tweaksToApply = @($tweakCheckboxes.GetEnumerator() | Where-Object { $_.Value.IsChecked -eq $true } | ForEach-Object { $_.Key })
    $appData       = $apps
    $tweakData     = $tweaks

    Write-UILog "=== TechFusion IT Bootstrap started ===" "#00b4d8"
    if ($doAdmin)  { Write-UILog "Admin  : $uname" "#e8f4fd" }
    if ($doApps)   { Write-UILog "Apps   : $($appsToInstall.Count) selected" "#e8f4fd" }
    if ($doTweaks) { Write-UILog "Tweaks : $($tweaksToApply.Count) selected" "#e8f4fd" }
    Write-UILog "---" "#1a2744"

    if (-not $doAdmin -and -not $doApps -and -not $doTweaks) {
        Write-UILog "Nothing selected - nothing to execute." "#e74c3c"
        $btnNext.Content = "Close"; $btnNext.IsEnabled = $true
        return
    }

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = "STA"
    $rs.ThreadOptions  = "ReuseThread"
    $rs.Open()

    $rs.SessionStateProxy.SetVariable("doAdmin",       $doAdmin)
    $rs.SessionStateProxy.SetVariable("uname",         $uname)
    $rs.SessionStateProxy.SetVariable("upwd",          $upwd)
    $rs.SessionStateProxy.SetVariable("doApps",        $doApps)
    $rs.SessionStateProxy.SetVariable("doTweaks",      $doTweaks)
    $rs.SessionStateProxy.SetVariable("appsToInstall", $appsToInstall)
    $rs.SessionStateProxy.SetVariable("tweaksToApply", $tweaksToApply)
    $rs.SessionStateProxy.SetVariable("appData",       $appData)
    $rs.SessionStateProxy.SetVariable("tweakData",     $tweakData)
    $rs.SessionStateProxy.SetVariable("appMode",       $script:appMode)
    $rs.SessionStateProxy.SetVariable("tweakMode",     $script:tweakMode)
    $rs.SessionStateProxy.SetVariable("window",        $window)
    $rs.SessionStateProxy.SetVariable("txtLog",        $txtLog)
    $rs.SessionStateProxy.SetVariable("logScroller",   $logScroller)
    $rs.SessionStateProxy.SetVariable("btnNext",       $btnNext)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs

    $workerCode = 'function Log{param([string]$M,[string]$K);if(-not $K){$K="#7a9bb5"};$ts=Get-Date -Format "HH:mm:ss";$window.Dispatcher.Invoke([Action]{$r=[System.Windows.Documents.Run]::new("["+$ts+"] "+$M+"`n");$r.Foreground=[System.Windows.Media.BrushConverter]::new().ConvertFromString($K);$txtLog.Inlines.Add($r);$logScroller.ScrollToBottom()})}
if($doAdmin){Log "--- Step 1: Admin Account ---" "#00b4d8";try{$sp=ConvertTo-SecureString $upwd -AsPlainText -Force;$ex=Get-LocalUser -Name $uname -ErrorAction SilentlyContinue;if($ex){Set-LocalUser -Name $uname -Password $sp;Log "  OK  Password updated for: $uname" "#22863a"}else{New-LocalUser -Name $uname -Password $sp -FullName $uname -Description "Local Administrator - TechFusion IT" -PasswordNeverExpires|Out-Null;Add-LocalGroupMember -Group "Administrators" -Member $uname|Out-Null;Log "  OK  Account created: $uname" "#22863a"}}catch{Log("  ERR "+$_.Exception.Message)"#e74c3c"}}
if($doApps){if($appMode -eq "install"){Log "--- Step 2: Installing Applications ---" "#00b4d8";foreach($id in $appsToInstall){$n=($appData|Where-Object{$_.Id -eq $id}|Select-Object -First 1).Name;Log "  Installing: $n" "#e8f4fd";try{winget install --id $id --silent --accept-package-agreements --accept-source-agreements 2>&1|Out-Null;if($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189){Log "  OK  $n" "#22863a"}else{Log "  ERR $n (exit code $LASTEXITCODE)" "#e74c3c"}}catch{Log("  ERR "+$_.Exception.Message)"#e74c3c"}}}else{Log "--- Step 2: Uninstalling Applications ---" "#c0392b";foreach($id in $appsToInstall){$n=($appData|Where-Object{$_.Id -eq $id}|Select-Object -First 1).Name;Log "  Uninstalling: $n" "#e8f4fd";try{winget uninstall --id $id --silent 2>&1|Out-Null;if($LASTEXITCODE -eq 0){Log "  OK  $n removed" "#22863a"}else{Log "  ERR $n (exit code $LASTEXITCODE)" "#e74c3c"}}catch{Log("  ERR "+$_.Exception.Message)"#e74c3c"}}}}
if($doTweaks){if($tweakMode -eq "apply"){Log "--- Step 3: Applying Tweaks ---" "#00b4d8"}else{Log "--- Step 3: Reverting Tweaks ---" "#c0392b"};foreach($id in $tweaksToApply){$n=($tweakData|Where-Object{$_.Id -eq $id}|Select-Object -First 1).Name;if($tweakMode -eq "apply"){Log "  Applying: $n" "#e8f4fd"}else{Log "  Reverting: $n" "#e8f4fd"};try{if($tweakMode -eq "apply"){switch($id){"DarkMode"{$p="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize";Set-ItemProperty $p -Name AppsUseLightTheme -Value 0 -Type DWord -Force;Set-ItemProperty $p -Name SystemUsesLightTheme -Value 0 -Type DWord -Force}"ShowFileExt"{Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name HideFileExt -Value 0 -Type DWord -Force}"ShowHidden"{Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name Hidden -Value 1 -Type DWord -Force}"DisableTelemetry"{$p="HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection";if(-not(Test-Path $p)){New-Item $p -Force|Out-Null};Set-ItemProperty $p -Name AllowTelemetry -Value 0 -Type DWord -Force}"ClassicContext"{$p="HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32";if(-not(Test-Path $p)){New-Item $p -Force|Out-Null};Set-ItemProperty $p -Name "(Default)" -Value "" -Force}"DisableBing"{Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name BingSearchEnabled -Value 0 -Type DWord -Force}"NumLock"{Set-ItemProperty "HKCU:\Control Panel\Keyboard" -Name InitialKeyboardIndicators -Value 2 -Force}"ClipboardHistory"{Set-ItemProperty "HKCU:\Software\Microsoft\Clipboard" -Name EnableClipboardHistory -Value 1 -Type DWord -Force}"DisableNewOutlook"{$p1="HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General";if(-not(Test-Path $p1)){New-Item $p1 -Force|Out-Null};Set-ItemProperty $p1 -Name HideNewOutlookToggle -Value 1 -Type DWord -Force;$p2="HKCU:\Software\Policies\Microsoft\office\16.0\outlook\preferences";if(-not(Test-Path $p2)){New-Item $p2 -Force|Out-Null};Set-ItemProperty $p2 -Name NewOutlookMigrationUserSetting -Value 0 -Type DWord -Force;Set-ItemProperty $p2 -Name NewOutlookAutomaticSetupUserSetting -Value 0 -Type DWord -Force;$p3="HKCU:\Software\Policies\Microsoft\office\16.0\outlook\options\general";if(-not(Test-Path $p3)){New-Item $p3 -Force|Out-Null};Set-ItemProperty $p3 -Name HideNewOutlookToggle -Value 1 -Type DWord -Force}}}else{switch($id){"DarkMode"{$p="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize";Set-ItemProperty $p -Name AppsUseLightTheme -Value 1 -Type DWord -Force;Set-ItemProperty $p -Name SystemUsesLightTheme -Value 1 -Type DWord -Force}"ShowFileExt"{Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name HideFileExt -Value 1 -Type DWord -Force}"ShowHidden"{Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name Hidden -Value 0 -Type DWord -Force}"DisableTelemetry"{Set-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name AllowTelemetry -Value 3 -Type DWord -Force}"ClassicContext"{$p="HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}";if(Test-Path $p){Remove-Item $p -Recurse -Force}}"DisableBing"{Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name BingSearchEnabled -Value 1 -Type DWord -Force}"NumLock"{Set-ItemProperty "HKCU:\Control Panel\Keyboard" -Name InitialKeyboardIndicators -Value 0 -Force}"ClipboardHistory"{Set-ItemProperty "HKCU:\Software\Microsoft\Clipboard" -Name EnableClipboardHistory -Value 0 -Type DWord -Force}"DisableNewOutlook"{$p1="HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General";Set-ItemProperty $p1 -Name HideNewOutlookToggle -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue;$p2="HKCU:\Software\Policies\Microsoft\office\16.0\outlook\preferences";Remove-ItemProperty $p2 -Name NewOutlookMigrationUserSetting -Force -ErrorAction SilentlyContinue;Remove-ItemProperty $p2 -Name NewOutlookAutomaticSetupUserSetting -Force -ErrorAction SilentlyContinue;$p3="HKCU:\Software\Policies\Microsoft\office\16.0\outlook\options\general";Set-ItemProperty $p3 -Name HideNewOutlookToggle -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue}}};Log "  OK  $n" "#22863a"}catch{Log("  ERR "+$_.Exception.Message)"#e74c3c"}}}
Log "---" "#1a2744"
Log "=== Bootstrap complete ===" "#00b4d8"
$window.Dispatcher.Invoke([Action]{$btnNext.Content="Close";$btnNext.IsEnabled=$true})'

    $ps.AddScript($workerCode) | Out-Null
    $ps.BeginInvoke()          | Out-Null
}

$btnNext.Add_Click({ if ($btnNext.Content -eq "Close") { $window.Close() } })

# ==============================================================================
#  LAUNCH
# ==============================================================================

# Exit the PowerShell process when the window is closed (X or Close button)
$window.Add_Closing({ [System.Environment]::Exit(0) })

Show-Step 1
$window.ShowDialog() | Out-Null
[System.Environment]::Exit(0)