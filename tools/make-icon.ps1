# Genera l'icona originale di PowerNotes (quadrato arrotondato blu d'accento con foglio e righe di testo). Nessuna risorsa esterna.
param([string]$Out = (Join-Path $PSScriptRoot '..\src\PowerNotes.App\Assets\powernotes.ico'))
Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null

function New-Frame([int]$n) {
    $bmp = New-Object System.Drawing.Bitmap $n, $n, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.Clear([System.Drawing.Color]::Transparent)
    $r = $n * 0.22
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $r * 2
    $path.AddArc(0, 0, $d, $d, 180, 90); $path.AddArc($n - $d, 0, $d, $d, 270, 90)
    $path.AddArc($n - $d, $n - $d, $d, $d, 0, 90); $path.AddArc(0, $n - $d, $d, $d, 90, 90); $path.CloseFigure()
    $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(79, 140, 255))), $path)
    # foglio
    $px = $n * 0.27; $py = $n * 0.16; $pw = $n * 0.46; $ph = $n * 0.68
    $g.FillRectangle([System.Drawing.Brushes]::White, $px, $py, $pw, $ph)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(79, 140, 255)), ([float][Math]::Max(1, $n * 0.04))
    foreach ($k in 0..3) { $y = $py + $ph * (0.22 + 0.19 * $k); $x2 = if ($k -eq 3) { $px + $pw * 0.55 } else { $px + $pw * 0.82 }; $g.DrawLine($pen, $px + $pw * 0.18, $y, $x2, $y) }
    $g.Dispose()
    return $bmp
}

$sizes = 16, 32, 48, 256
$ms = New-Object System.IO.MemoryStream
$bw = New-Object System.IO.BinaryWriter $ms
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
$blobs = @()
foreach ($s in $sizes) { $b = New-Frame $s; $m = New-Object System.IO.MemoryStream; $b.Save($m, [System.Drawing.Imaging.ImageFormat]::Png); $blobs += , $m.ToArray(); $b.Dispose() }
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $s = $sizes[$i]; $w = if ($s -ge 256) { 0 } else { $s }
    $bw.Write([byte]$w); $bw.Write([byte]$w); $bw.Write([byte]0); $bw.Write([byte]0); $bw.Write([uint16]1); $bw.Write([uint16]32)
    $bw.Write([uint32]$blobs[$i].Length); $bw.Write([uint32]$offset); $offset += $blobs[$i].Length
}
foreach ($b in $blobs) { $bw.Write($b) }
[IO.File]::WriteAllBytes((Resolve-Path -LiteralPath (Split-Path $Out)).Path + '\' + (Split-Path $Out -Leaf), $ms.ToArray())
