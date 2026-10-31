# Helper COM Word: avvio isolato, watchdog per PID, chiusura garantita. Uso: . .\WordCom.ps1
function Start-PNWord([int]$TimeoutSec = 240) {
    $pre = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $w = New-Object -ComObject Word.Application
    $procId = 0
    for ($i = 0; $i -lt 40 -and -not $procId; $i++) {
        Start-Sleep -Milliseconds 250
        $procId = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $pre -notcontains $_.Id } | ForEach-Object { $_.Id }) | Select-Object -First 1
    }
    if (-not $procId) { try { $w.Quit(0) } catch {}; throw 'PID di Word non identificato' }
    $w.Visible = $false
    $w.DisplayAlerts = 0
    $w.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: nessuna macro
    # watchdog: processo separato che uccide SOLO il PID di questa istanza
    $wd = Start-Process powershell -WindowStyle Hidden -PassThru -ArgumentList "-NoProfile", "-Command", "Start-Sleep $TimeoutSec; Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue"
    [pscustomobject]@{ App = $w; Pid = $procId; Watchdog = $wd }
}

function Stop-PNWord($s) {
    try { $s.App.Quit(0) } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($s.App) } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 800
    if (Get-Process -Id $s.Pid -ErrorAction SilentlyContinue) { Stop-Process -Id $s.Pid -Force -ErrorAction SilentlyContinue }
    try { Stop-Process -Id $s.Watchdog.Id -Force -ErrorAction SilentlyContinue } catch {}
}
