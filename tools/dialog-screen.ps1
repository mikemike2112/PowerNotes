# Cattura dello SCHERMO VERO (CopyFromScreen) della finestra di salvataggio, senza input: l'app apre il pannello con --dialog-hold,
# lo script legge il rettangolo via UI Automation, cattura e chiude il processo per PID. Uso: dialog-screen.ps1 -Exe <exe> -Out <png> [-Light]
param([Parameter(Mandatory)][string]$Exe, [Parameter(Mandatory)][string]$Out, [switch]$Light)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.Windows.Forms
Add-Type -Namespace PN -Name D -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [System.Runtime.InteropServices.DllImport("user32.dll")] public static extern int GetDpiForSystem();'
[void][PN.D]::SetProcessDPIAware()
$env:POWERNOTES_SETTINGS = Join-Path $env:TEMP 'pn-dlg-settings.json'; Remove-Item $env:POWERNOTES_SETTINGS -ea SilentlyContinue
$env:POWERNOTES_DEFAULT_FOLDER = 'C:\Documenti'
$a = @('--dialog-hold'); if ($Light) { $a += '--light' }
$p = Start-Process $Exe -ArgumentList $a -PassThru
try {
    $A = [System.Windows.Automation.AutomationElement]; $win = $null
    $c = [System.Windows.Automation.PropertyCondition]::new($A::NameProperty, [string]'Salva')
    for ($i = 0; $i -lt 60 -and -not $win; $i++) {
        Start-Sleep -Milliseconds 300
        $btn = $A::RootElement.FindFirst('Descendants', [System.Windows.Automation.AndCondition]::new($c, [System.Windows.Automation.PropertyCondition]::new($A::ProcessIdProperty, [int]$p.Id)))
        if ($btn) { $win = $btn }
    }
    if (-not $win) { throw 'finestra di salvataggio non trovata' }
    Start-Sleep -Milliseconds 800
    $w = $win; while ($w -and $w.Current.ControlType.ProgrammaticName -ne 'ControlType.Window') { $w = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($w) }
    $r = $w.Current.BoundingRectangle
    $bmp = New-Object System.Drawing.Bitmap ([int]$r.Width), ([int]$r.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen([int]$r.X, [int]$r.Y, 0, 0, $bmp.Size); $g.Dispose()
    $bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    "dpi=$([PN.D]::GetDpiForSystem()) rect=$r"
} finally { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
