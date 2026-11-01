# Esporta in PDF con il vero Word (COM) i documenti indicati. Uso: word-pdf.ps1 -Docs a.docx,b.docx -OutDir cartella
param([Parameter(Mandatory)][string[]]$Docs, [Parameter(Mandatory)][string]$OutDir, [int]$Timeout = 300)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WordCom.ps1')
$Docs = @($Docs | ForEach-Object { $_ -split ',' })
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path
$s0 = Start-PNWord $Timeout
try {
    foreach ($p in $Docs) {
        $full = (Resolve-Path $p).Path
        $d = $w = $null
        try {
            $d = $s0.App.Documents.Open($full, [ref]$false, [ref]$true, [ref]$false)
            $out = Join-Path $OutDir ([IO.Path]::GetFileNameWithoutExtension($full) + '.word.pdf')
            $d.SaveAs2([string]$out, [int]17)
        } finally { if ($d) { try { $d.Close(0) } catch {} } }
    }
} finally { Stop-PNWord $s0 }
