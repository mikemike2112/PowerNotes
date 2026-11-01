# Prova con input REALE (mouse) della scheda Powers: tema, persistenza dopo riavvio, file impostazioni corrotto.
# Avvia PowerNotes.exe con impostazioni isolate (POWERNOTES_SETTINGS), clicca con il mouse vero, fa screenshot della sola finestra,
# chiude l'app per PID proprio. Uso: ui-powers-test.ps1 -Exe <PowerNotes.exe> -OutDir <cartella> -OutJson <file.json>
param([Parameter(Mandatory)][string]$Exe, [Parameter(Mandatory)][string]$OutDir, [Parameter(Mandatory)][string]$OutJson)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.Windows.Forms
Add-Type -Namespace PN -Name Native -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool SetWindowPos(System.IntPtr h, System.IntPtr after, int x, int y, int cx, int cy, uint f);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern System.IntPtr WindowFromPoint(int x, int y);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern System.IntPtr GetAncestor(System.IntPtr h, uint f);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool PrintWindow(System.IntPtr h, System.IntPtr hdc, uint flags);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, System.UIntPtr e);
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h);
'@
[void][PN.Native]::SetProcessDPIAware()
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path
$settings = Join-Path $OutDir 'settings.json'
$env:POWERNOTES_SETTINGS = $settings
$A = [System.Windows.Automation.AutomationElement]
$res = [ordered]@{}
trap { $res | ConvertTo-Json | Write-Host; Write-Error ("riga " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message); break }

function Start-App {
    $p = Start-Process $Exe -PassThru
    $root = $A::RootElement; $win = $null
    for ($i = 0; $i -lt 60 -and -not $win; $i++) {
        Start-Sleep -Milliseconds 500
        $win = $root.FindFirst('Children', [System.Windows.Automation.PropertyCondition]::new($A::ProcessIdProperty, [int]$p.Id))
    }
    if (-not $win) { throw 'finestra non trovata' }
    if (-not (Try-Find $win 'Powers' 20000)) { throw 'interfaccia non pronta' }
    [void][PN.Native]::SetForegroundWindow([IntPtr]$win.Current.NativeWindowHandle)
    Start-Sleep -Milliseconds 500
    $script:topHwnd = $win.Current.NativeWindowHandle
    [pscustomobject]@{ Proc = $p; Win = $win }
}
function Stop-App($app) {
    try { [void]$app.Proc.CloseMainWindow() } catch {}
    if (-not $app.Proc.WaitForExit(5000)) { Stop-Process -Id $app.Proc.Id -Force }
}
function Try-Find($win, $name, $ms = 1500) {
    $end = [DateTime]::Now.AddMilliseconds($ms)
    while ([DateTime]::Now -lt $end) { $e = $win.FindFirst('Descendants', [System.Windows.Automation.PropertyCondition]::new($A::NameProperty, [string]$name)); if ($e) { return $e }; Start-Sleep -Milliseconds 150 }
    return $null
}
# Il desktop può essere condiviso con altre finestre: porta in primo piano la nostra e ripeti il clic finché l'effetto atteso compare.
function Open-Powers($app) {
    for ($k = 0; $k -lt 6; $k++) {
        [void][PN.Native]::SetForegroundWindow([IntPtr]$app.Win.Current.NativeWindowHandle); Start-Sleep -Milliseconds 300
        Click-Real (Try-Find $app.Win 'Powers' 3000)
        $sw = Try-Find $app.Win 'Modalità scura' 1500
        if ($sw) { return $sw }
    }
    throw 'scheda Powers non aperta'
}
function Find-ByName($win, $name) {
    for ($i = 0; $i -lt 20; $i++) { $e = $win.FindFirst('Descendants', [System.Windows.Automation.PropertyCondition]::new($A::NameProperty, [string]$name)); if ($e) { return $e }; Start-Sleep -Milliseconds 250 }
    throw "elemento non trovato: $name"
}
function Click-Real($el) {
    [void][PN.Native]::SetWindowPos([IntPtr]$script:topHwnd, [IntPtr](-1), 0, 0, 0, 0, 0x13)   # HWND_TOPMOST, senza spostare né attivare
    $r = $el.Current.BoundingRectangle
    [void][PN.Native]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
    Start-Sleep -Milliseconds 200
    [PN.Native]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 80; [PN.Native]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 700
}
function Shot($win, $path) {
    # PrintWindow sulla SOLA finestra dell'app (altre finestre sul desktop non compaiono)
    $r = $win.Current.BoundingRectangle
    $bmp = New-Object System.Drawing.Bitmap ([int]$r.Width), ([int]$r.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc()
    [void][PN.Native]::PrintWindow([IntPtr]$win.Current.NativeWindowHandle, $hdc, 2)
    $g.ReleaseHdc($hdc); $g.Dispose()
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
}
function Wait-State($win, $want, $ms = 4000) {
    $end = [DateTime]::Now.AddMilliseconds($ms); $s = $null
    while ([DateTime]::Now -lt $end) { $e = Try-Find $win 'Modalità scura' 500; if ($e) { $s = Toggle-State $e; if ($s -eq $want) { return $s } }; Start-Sleep -Milliseconds 100 }
    return $s
}
function Toggle-State($el) { $el.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Current.ToggleState.ToString() }

# 1) prima apertura: nessun file di impostazioni -> scuro; clic reale su Powers e sull'interruttore
if (Test-Path $settings) { Remove-Item $settings }
$app = Start-App
try {
    $sw = Open-Powers $app
    $res.firstRunState = Toggle-State $sw
    Shot $app.Win (Join-Path $OutDir 'powers-scuro.png')
    for ($k = 0; $k -lt 6; $k++) {
        [void][PN.Native]::SetForegroundWindow([IntPtr]$app.Win.Current.NativeWindowHandle); Start-Sleep -Milliseconds 300
        $cur = Try-Find $app.Win 'Modalità scura' 3000
        if ($cur -and (Toggle-State $cur) -eq 'Off') { break }
        if ($cur) { Click-Real $cur }
    }
    $res.afterClickState = Wait-State $app.Win 'Off'
    Shot $app.Win (Join-Path $OutDir 'powers-chiaro.png')
} finally { Stop-App $app }
$res.savedJson = if (Test-Path $settings) { [IO.File]::ReadAllText($settings) } else { '' }

# 2) riavvio: lo stato chiaro deve persistere
$app = Start-App
try {
    $res.restartState = Toggle-State (Open-Powers $app)
    Shot $app.Win (Join-Path $OutDir 'riavvio-chiaro.png')
} finally { Stop-App $app }

# 3) file corrotto: nessun crash, si parte scuro
[IO.File]::WriteAllText($settings, '{ questo non è json ::::')
$app = Start-App
try {
    $res.corruptAlive = -not $app.Proc.HasExited
    $res.corruptState = Toggle-State (Open-Powers $app)
} finally { Stop-App $app }
[IO.File]::WriteAllText($OutJson, ($res | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
