# Inventario del ribbon di Word con UI Automation: apre Word (finestra visibile) su un documento vuoto creato qui,
# legge schede, gruppi e comandi, scrive un JSON. Nessun documento dell'utente viene aperto.
param([string]$Out = (Join-Path $PSScriptRoot '..\docs\word-ribbon-inventory.json'), [int]$Timeout = 400)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WordCom.ps1')
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$A = [System.Windows.Automation.AutomationElement]
$CT = [System.Windows.Automation.ControlType]
$s0 = Start-PNWord $Timeout
$res = [ordered]@{}
try {
    $w = $s0.App
    $w.Visible = $true
    $d = $w.Documents.Add()
    Start-Sleep 3
    $root = $A::RootElement
    $pidCond = New-Object System.Windows.Automation.PropertyCondition ($A::ProcessIdProperty), ([int]$s0.Pid)
    $win = $null
    for ($i = 0; $i -lt 20 -and -not $win; $i++) { Start-Sleep -Milliseconds 500; $win = $root.FindFirst('Children', $pidCond) }
    if (-not $win) { throw 'finestra Word non trovata' }
    $tabs = $win.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition ($A::ControlTypeProperty), $CT::TabItem))
    $names = @(); foreach ($t in $tabs) { $names += $t.Current.Name }
    $res['schede'] = $names
    foreach ($t in $tabs) {
        $n = $t.Current.Name
        if (-not $n) { continue }
        try { $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() } catch { continue }
        Start-Sleep -Milliseconds 700
        $groups = [ordered]@{}
        $gs = $win.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition ($A::ControlTypeProperty), $CT::Group))
        foreach ($g in $gs) {
            $gn = $g.Current.Name
            if (-not $gn -or $g.Current.IsOffscreen) { continue }
            $items = @()
            foreach ($ct in @($CT::Button, $CT::SplitButton, $CT::ComboBox, $CT::CheckBox, $CT::MenuItem, $CT::Edit)) {
                $els = $g.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition ($A::ControlTypeProperty), $ct))
                foreach ($e in $els) { if ($e.Current.Name -and -not $e.Current.IsOffscreen) { $items += $e.Current.Name } }
            }
            if ($items.Count) { $groups[$gn] = @($items | Select-Object -Unique) }
        }
        $res[$n] = $groups
    }
    $d.Close(0)
} finally { Stop-PNWord $s0 }
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
[IO.File]::WriteAllText($Out, ($res | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
