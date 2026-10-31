# Apre file .docx in Word (COM, sola lettura, finestra nascosta) e scrive un JSON con esito, pagine e contenuto.
# Un file che richiede riparazione o non si apre produce ok=false. Uso: word-inspect.ps1 -Paths a.docx,b.docx -Out risultato.json [-Detail]
param([Parameter(Mandatory)][string[]]$Paths, [Parameter(Mandatory)][string]$Out, [switch]$Detail, [int]$Timeout = 300)
$ErrorActionPreference = 'Stop'
$Paths = @($Paths | ForEach-Object { $_ -split ',' })
. (Join-Path $PSScriptRoot 'WordCom.ps1')
$s0 = Start-PNWord $Timeout
$res = [ordered]@{}
try {
    $w = $s0.App
    foreach ($p in $Paths) {
        $full = (Resolve-Path $p).Path
        $r = [ordered]@{ ok = $false; error = $null }
        $d = $null
        try {
            $d = $w.Documents.Open($full, [ref]$false, [ref]$true, [ref]$false)
            $d.Repaginate()
            $r.pages = $d.ComputeStatistics(2)
            $r.text = $d.Content.Text
            $r.paragraphs = $d.Paragraphs.Count
            $r.tables = $d.Tables.Count
            $r.inlineShapes = $d.InlineShapes.Count
            $r.shapes = $d.Shapes.Count
            $r.sections = $d.Sections.Count
            $r.windowTitle = $w.Caption
            $r.compat = $d.CompatibilityMode
            $bp = $d.BuiltInDocumentProperties
            $it = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $bp, @('Author'))
            $r.author = [System.__ComObject].InvokeMember('Value', 'GetProperty', $null, $it, $null)
            $it = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $bp, @('Last Author'))
            $r.lastAuthor = [System.__ComObject].InvokeMember('Value', 'GetProperty', $null, $it, $null)
            if ($Detail) {
                $lines = @()
                $n = $d.Paragraphs.Count
                for ($i = 1; $i -le $n; $i++) {
                    $rg = $d.Paragraphs.Item($i).Range
                    $rg.Collapse(1)
                    $lines += ,@([int]$rg.Information(3), [math]::Round([double]$rg.Information(6), 2), [math]::Round([double]$rg.Information(5), 2))
                }
                $r.lines = $lines
            }
            $r.ok = $true
        } catch { $r.error = $_.Exception.Message }
        finally { if ($d) { try { $d.Close(0) } catch {} } }
        $res[$full] = $r
    }
} finally { Stop-PNWord $s0 }
[IO.File]::WriteAllText($Out, ($res | ConvertTo-Json -Depth 6 -Compress), (New-Object Text.UTF8Encoding($false)))
