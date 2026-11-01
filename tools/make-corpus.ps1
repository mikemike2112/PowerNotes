# Rigenera il corpus tests/corpus/office-made/*.docx con il vero Word (COM). Nessun dato personale: Author/LastAuthor = "Corpus".
param([string]$Out = (Join-Path $PSScriptRoot '..\tests\corpus\office-made'), [string]$Only = '')
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WordCom.ps1')
$Out = (Resolve-Path $Out).Path
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

$sentences = @(
 'Il rapido volpino marrone salta sopra il cane pigro mentre il sole tramonta lentamente dietro le colline.',
 'La stesura di un documento richiede attenzione ai dettagli, alla struttura e soprattutto alla leggibilità del testo.',
 'Ogni paragrafo dovrebbe sviluppare una sola idea, collegandola con naturalezza a quella che la precede e a quella che la segue.',
 'Le tabelle organizzano i dati in righe e colonne, permettendo al lettore di confrontare rapidamente valori diversi.',
 'Una buona impaginazione guida lo sguardo, lascia respirare il testo e rende piacevole la lettura anche dei documenti più lunghi.')
function Lorem([int]$n, [int]$seed = 0) { (0..($n-1) | % { $sentences[($_ + $seed) % $sentences.Count] }) -join ' ' }

# immagine di prova generata qui (originale, nessuna risorsa esterna)
$png = Join-Path $env:TEMP 'pn-corpus-img.png'
$bmp = New-Object System.Drawing.Bitmap 320, 180
$g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::FromArgb(79,140,255))
$g.FillEllipse([System.Drawing.Brushes]::White, 110, 40, 100, 100); $g.Dispose(); $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()

function P($sel, $text) { $sel.TypeText($text); $sel.TypeParagraph() }
function Style($doc, $id) { $doc.Styles.Item([int]$id) }

$sc = [ordered]@{}
$sc['plain-paragraphs'] = { param($w,$d,$s) 1..8 | % { P $s (Lorem 5 $_) } }
$sc['text-formatting'] = { param($w,$d,$s)
  $s.Font.Bold=1; $s.TypeText('Grassetto '); $s.Font.Bold=0
  $s.Font.Italic=1; $s.TypeText('Corsivo '); $s.Font.Italic=0
  $s.Font.Underline=1; $s.TypeText('Sottolineato '); $s.Font.Underline=0
  $s.Font.StrikeThrough=1; $s.TypeText('Barrato '); $s.Font.StrikeThrough=0
  $s.Font.Color=255; $s.TypeText('Rosso '); $s.Font.Color=-16777216
  $s.Font.Name='Times New Roman'; $s.TypeText('Times '); $s.Font.Name='Courier New'; $s.TypeText('Courier '); $s.Font.Name='Arial'; $s.TypeText('Arial ')
  $s.Font.Size=20; $s.TypeText('Grande '); $s.Font.Size=8; $s.TypeText('Piccolo '); $s.Font.Size=11
  $s.Range.HighlightColorIndex=7; $s.TypeText('Evidenziato '); $s.Range.HighlightColorIndex=0
  $s.TypeParagraph(); P $s (Lorem 3) }
$sc['font-effects'] = { param($w,$d,$s)
  $s.Font.SmallCaps=1; $s.TypeText('Maiuscoletto '); $s.Font.SmallCaps=0
  $s.Font.AllCaps=1; $s.TypeText('Tutto maiuscolo '); $s.Font.AllCaps=0
  $s.TypeText('H'); $s.Font.Subscript=1; $s.TypeText('2'); $s.Font.Subscript=0; $s.TypeText('O e x'); $s.Font.Superscript=1; $s.TypeText('2'); $s.Font.Superscript=0
  $s.TypeParagraph() }
$sc['alignment'] = { param($w,$d,$s)
  0..3 | % { $s.ParagraphFormat.Alignment = [int]$_; P $s (Lorem 4 $_) } }
$sc['spacing-indents'] = { param($w,$d,$s)
  $s.ParagraphFormat.SpaceBefore=12; $s.ParagraphFormat.SpaceAfter=18; P $s (Lorem 3)
  $s.ParagraphFormat.SpaceBefore=0; $s.ParagraphFormat.SpaceAfter=6
  $s.ParagraphFormat.LineSpacingRule=1; P $s (Lorem 4 1)
  $s.ParagraphFormat.LineSpacingRule=2; P $s (Lorem 4 2)
  $s.ParagraphFormat.LineSpacingRule=4; $s.ParagraphFormat.LineSpacing=20; P $s (Lorem 4 3)
  $s.ParagraphFormat.LineSpacingRule=0
  $s.ParagraphFormat.LeftIndent=36; $s.ParagraphFormat.RightIndent=36; P $s (Lorem 4 4)
  $s.ParagraphFormat.LeftIndent=0; $s.ParagraphFormat.RightIndent=0; $s.ParagraphFormat.FirstLineIndent=28; P $s (Lorem 4)
  $s.ParagraphFormat.LeftIndent=40; $s.ParagraphFormat.FirstLineIndent=-40; P $s (Lorem 4 2) }
$sc['styles-headings'] = { param($w,$d,$s)
  $s.Style=(Style $d -63); P $s 'Titolo del documento'
  $s.Style=(Style $d -75); P $s 'Sottotitolo'
  $s.Style=(Style $d -2); P $s 'Capitolo uno'
  $s.Style=(Style $d -1); P $s (Lorem 3)
  $s.Style=(Style $d -3); P $s 'Sezione 1.1'
  $s.Style=(Style $d -1); P $s (Lorem 3 1)
  $s.Style=(Style $d -4); P $s 'Sottosezione 1.1.1'
  $s.Style=(Style $d -1); P $s (Lorem 3 2)
  $s.Style=(Style $d -181); P $s 'Una citazione ben messa.'
  $s.Style=(Style $d -1) }
$sc['style-custom'] = { param($w,$d,$s)
  $st=$d.Styles.Add('Evidenza',1); $st.Font.Name='Georgia'; $st.Font.Size=14; $st.Font.Bold=1; $st.Font.Color=16711680; $st.ParagraphFormat.SpaceAfter=10; $st.ParagraphFormat.Alignment=1
  $s.Style=$st; P $s 'Paragrafo con stile personalizzato'; $s.Style=(Style $d -1); P $s (Lorem 3) }
$sc['list-bullets'] = { param($w,$d,$s)
  P $s 'Elenco puntato:'; $s.Range.ListFormat.ApplyBulletDefault(); 1..4 | % { P $s "Voce numero $_" }; $s.Range.ListFormat.RemoveNumbers(); P $s 'Fine.' }
$sc['list-numbered'] = { param($w,$d,$s)
  P $s 'Elenco numerato:'; $s.Range.ListFormat.ApplyNumberDefault(); 1..5 | % { P $s "Passo $_ della procedura" }; $s.Range.ListFormat.RemoveNumbers(); P $s 'Fine.' }
$sc['list-multilevel'] = { param($w,$d,$s)
  $lt=$w.ListGalleries.Item(3).ListTemplates.Item(1)
  $s.Range.ListFormat.ApplyListTemplate($lt)
  P $s 'Livello 1 A'; $s.Range.ListFormat.ListIndent(); P $s 'Livello 2 A'; P $s 'Livello 2 B'; $s.Range.ListFormat.ListIndent(); P $s 'Livello 3 A'
  $s.Range.ListFormat.ListOutdent(); $s.Range.ListFormat.ListOutdent(); P $s 'Livello 1 B'; $s.Range.ListFormat.RemoveNumbers() }
$sc['table-simple'] = { param($w,$d,$s)
  P $s 'Prima della tabella'; $t=$d.Tables.Add($s.Range,4,3); $t.Borders.Enable=$true
  1..4 | % { $r=$_; 1..3 | % { $t.Cell($r,[int]$_).Range.Text="R$r C$_" } }
  $s.EndKey(6) | Out-Null; P $s 'Dopo la tabella' }
$sc['table-merged-cells'] = { param($w,$d,$s)
  $t=$d.Tables.Add($s.Range,4,4); $t.Borders.Enable=$true
  1..4 | % { $r=$_; 1..4 | % { $t.Cell($r,[int]$_).Range.Text="R$r C$_" } }
  $t.Cell(1,1).Merge($t.Cell(1,3)); $t.Cell(2,4).Merge($t.Cell(4,4)); $t.Cell(3,1).Merge($t.Cell(4,2))
  $s.EndKey(6) | Out-Null; P $s 'Dopo la tabella con celle unite' }
$sc['table-borders-shading'] = { param($w,$d,$s)
  $t=$d.Tables.Add($s.Range,3,3); $t.Borders.Enable=$true; $t.Borders.OutsideLineWidth=24; $t.Borders.InsideLineStyle=1
  1..3 | % { $r=$_; 1..3 | % { $t.Cell($r,[int]$_).Range.Text="Cella $r,$_" } }
  1..3 | % { $t.Cell(1,[int]$_).Shading.BackgroundPatternColor=15849925; $t.Cell(1,[int]$_).Range.Font.Bold=1 }
  $t.Cell(2,2).Shading.BackgroundPatternColor=13421823; $t.Cell(2,3).Range.ParagraphFormat.Alignment=2
  $s.EndKey(6) | Out-Null; P $s '' }
$sc['table-style-header-repeat'] = { param($w,$d,$s)
  $t=$d.Tables.Add($s.Range,70,3); try { $t.Style='Tabella griglia 4 - colore 1' } catch { $t.Borders.Enable=$true }
  1..70 | % { $r=$_; 1..3 | % { $t.Cell($r,[int]$_).Range.Text="R$r C$_" } }
  $t.Rows.Item(1).HeadingFormat=-1
  $s.EndKey(6) | Out-Null; P $s 'Fine' }
$sc['table-nested'] = { param($w,$d,$s)
  $t=$d.Tables.Add($s.Range,2,2); $t.Borders.Enable=$true
  $t.Cell(1,1).Range.Text='Esterna 1'; $t.Cell(1,2).Range.Text='Esterna 2'; $t.Cell(2,1).Range.Text='Esterna 3'
  $n=$d.Tables.Add($t.Cell(2,2).Range,2,2); $n.Borders.Enable=$true; $n.Cell(1,1).Range.Text='in 1'; $n.Cell(2,2).Range.Text='in 4'
  $s.EndKey(6) | Out-Null; P $s '' }
$sc['table-multipage'] = { param($w,$d,$s)
  $t=$d.Tables.Add($s.Range,60,3); $t.Borders.Enable=$true
  1..60 | % { $r=$_; 1..3 | % { $t.Cell($r,[int]$_).Range.Text="Riga $r col $_" } }
  $s.EndKey(6) | Out-Null; P $s 'Fine' }
$sc['image-inline'] = { param($w,$d,$s)
  P $s 'Immagine inline:'; $d.InlineShapes.AddPicture($png,$false,$true,$s.Range) | Out-Null; $s.EndKey(5) | Out-Null; $s.TypeParagraph(); P $s (Lorem 2)
  $sh=$d.InlineShapes.AddPicture($png,$false,$true,$s.Range); $sh.LockAspectRatio=-1; $sh.Width=100; $s.EndKey(5)|Out-Null; $s.TypeText(' testo accanto all immagine piccola.'); $s.TypeParagraph() }
$sc['page-setup-landscape'] = { param($w,$d,$s)
  $d.PageSetup.Orientation=1; P $s 'Pagina in orizzontale'; 1..4 | % { P $s (Lorem 5 $_) } }
$sc['page-setup-letter'] = { param($w,$d,$s)
  $d.PageSetup.PaperSize=2; $d.PageSetup.LeftMargin=54; $d.PageSetup.RightMargin=54; P $s 'Formato Letter margini 0,75 pollici'; 1..4 | % { P $s (Lorem 5 $_) } }
$sc['page-margins-narrow'] = { param($w,$d,$s)
  $ps=$d.PageSetup; $ps.TopMargin=36;$ps.BottomMargin=36;$ps.LeftMargin=36;$ps.RightMargin=36; 1..10 | % { P $s (Lorem 5 $_) } }
$sc['page-break'] = { param($w,$d,$s)
  P $s 'Pagina uno'; $s.InsertBreak(7); P $s 'Pagina due'; $s.InsertBreak(7); P $s 'Pagina tre' }
$sc['page-break-before-keep'] = { param($w,$d,$s)
  1..6 | % { P $s (Lorem 5 $_) }
  $s.ParagraphFormat.PageBreakBefore=-1; P $s 'Paragrafo con interruzione prima'; $s.ParagraphFormat.PageBreakBefore=0
  $s.ParagraphFormat.KeepWithNext=-1; P $s 'Titolo che resta con il successivo'; $s.ParagraphFormat.KeepWithNext=0; P $s (Lorem 3) }
$sc['header-footer'] = { param($w,$d,$s)
  $d.Sections.Item(1).Headers.Item(1).Range.Text='Intestazione del documento'
  $d.Sections.Item(1).Footers.Item(1).Range.Text='Piè di pagina'
  1..10 | % { P $s (Lorem 5 $_) } }
$sc['page-numbers'] = { param($w,$d,$s)
  $f=$d.Sections.Item(1).Footers.Item(1).Range; $f.Text='Pagina '; $f.ParagraphFormat.Alignment=1
  $r=$d.Sections.Item(1).Footers.Item(1).Range; $r.Collapse(0); $d.Fields.Add($r,33) | Out-Null
  $r=$d.Sections.Item(1).Footers.Item(1).Range; $r.InsertAfter(' di '); $r=$d.Sections.Item(1).Footers.Item(1).Range; $r.Collapse(0); $d.Fields.Add($r,26) | Out-Null
  1..16 | % { P $s (Lorem 5 $_) } }
$sc['section-break'] = { param($w,$d,$s)
  P $s 'Sezione uno verticale'; $s.InsertBreak(2); $d.Sections.Item(2).PageSetup.Orientation=1; P $s 'Sezione due orizzontale'; $s.InsertBreak(2); $d.Sections.Item(3).PageSetup.Orientation=0; P $s 'Sezione tre' }
$sc['tabs'] = { param($w,$d,$s)
  $s.ParagraphFormat.TabStops.ClearAll(); $s.ParagraphFormat.TabStops.Add(150,0)|Out-Null; $s.ParagraphFormat.TabStops.Add(300,2)|Out-Null
  P $s "Nome`tCognome`tEtà"; P $s "Anna`tRossi`t30"; P $s "Luigi`tVerdi`t41"; $s.ParagraphFormat.TabStops.ClearAll(); P $s "Tab`tdefault`tqui" }
$sc['line-breaks'] = { param($w,$d,$s)
  $s.TypeText('Prima riga'); $s.InsertBreak(6); $s.TypeText('Seconda riga dopo interruzione'); $s.TypeParagraph(); P $s (Lorem 2) }
$sc['font-sizes'] = { param($w,$d,$s)
  8,9,10,11,12,14,16,18,24,36 | % { $s.Font.Size=[int]$_; P $s "Corpo $_ pt - testo di prova" } }
$sc['hyperlink'] = { param($w,$d,$s)
  $s.TypeText('Vai a '); $d.Hyperlinks.Add($s.Range,'https://example.org/','','','esempio') | Out-Null; $s.EndKey(5)|Out-Null; $s.TypeParagraph() }
$sc['footnote'] = { param($w,$d,$s)
  $s.TypeText('Testo con nota'); $d.Footnotes.Add($s.Range,[Type]::Missing,'Testo della nota a piè di pagina.') | Out-Null; $s.EndKey(6)|Out-Null; P $s '.' }
$sc['columns-two'] = { param($w,$d,$s)
  $d.PageSetup.TextColumns.SetCount(2); 1..8 | % { P $s (Lorem 5 $_) } }
$sc['track-changes'] = { param($w,$d,$s)
  P $s 'Testo originale prima della revisione.'; $d.TrackRevisions=$true; $s.TypeText('Testo inserito in revisione.'); $s.TypeParagraph(); $d.TrackRevisions=$false }
$sc['shape-textbox'] = { param($w,$d,$s)
  P $s 'Prima della casella'; $d.Shapes.AddTextbox(1,100,200,200,60,$s.Range) | Out-Null; $d.Shapes.Item(1).TextFrame.TextRange.Text='Testo nella casella'; P $s 'Dopo' }
$sc['long-document'] = { param($w,$d,$s)
  1..12 | % { $s.Style=(Style $d -2); P $s "Capitolo $_"; $s.Style=(Style $d -1); 1..4 | % { P $s (Lorem 6 $_) } } }
$sc['mixed-report'] = { param($w,$d,$s)
  $s.Style=(Style $d -63); P $s 'Relazione trimestrale'; $s.Style=(Style $d -1); P $s (Lorem 3)
  $s.Style=(Style $d -2); P $s 'Risultati'; $s.Style=(Style $d -1); P $s (Lorem 4 1)
  $t=$d.Tables.Add($s.Range,4,3); $t.Borders.Enable=$true; 1..4 | % { $r=$_; 1..3 | % { $t.Cell($r,[int]$_).Range.Text="V$r$_" } }; $t.Rows.Item(1).Range.Font.Bold=1; $t.Rows.Item(1).Shading.BackgroundPatternColor=14277081
  $s.EndKey(6)|Out-Null; $s.Style=(Style $d -3); P $s 'Punti chiave'; $s.Style=(Style $d -1)
  $s.Range.ListFormat.ApplyBulletDefault(); 1..3 | % { P $s "Punto $_" }; $s.Range.ListFormat.RemoveNumbers()
  P $s (Lorem 6 2); 1..6 | % { P $s (Lorem 6 $_) }
  $f=$d.Sections.Item(1).Footers.Item(1).Range; $f.Text='Pagina '; $r=$d.Sections.Item(1).Footers.Item(1).Range; $r.Collapse(0); $d.Fields.Add($r,33)|Out-Null }
$sc['mixed-letter'] = { param($w,$d,$s)
  $s.ParagraphFormat.Alignment=2; P $s 'Roma, 1 gennaio 2026'; $s.ParagraphFormat.Alignment=0
  P $s 'Gentile cliente,'; 1..3 | % { $s.ParagraphFormat.Alignment=3; P $s (Lorem 4 $_) }; $s.ParagraphFormat.Alignment=0; P $s 'Cordiali saluti,'; $s.Font.Bold=1; P $s 'L ufficio'; $s.Font.Bold=0 }
$sc['mixed-illustrated'] = { param($w,$d,$s)
  $s.Style=(Style $d -2); P $s 'Documento illustrato'; $s.Style=(Style $d -1); P $s (Lorem 4)
  $d.InlineShapes.AddPicture($png,$false,$true,$s.Range)|Out-Null; $s.EndKey(5)|Out-Null; $s.ParagraphFormat.Alignment=1; $s.TypeParagraph(); $s.ParagraphFormat.Alignment=0
  P $s (Lorem 6 2); $s.Range.ListFormat.ApplyNumberDefault(); 1..3 | % { P $s "Elemento $_" }; $s.Range.ListFormat.RemoveNumbers(); 1..8 | % { P $s (Lorem 5 $_) } }

function Patch-Props([string]$path) {
  $zip = [System.IO.Compression.ZipFile]::Open($path, 'Update')
  try {
    foreach ($n in 'docProps/core.xml','docProps/app.xml') {
      $e = $zip.GetEntry($n); if (-not $e) { continue }
      $sr = New-Object IO.StreamReader($e.Open()); $x = $sr.ReadToEnd(); $sr.Close()
      $x = [regex]::Replace($x, '(<dc:creator>)[^<]*(</dc:creator>)', '${1}Corpus${2}')
      $x = [regex]::Replace($x, '(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)', '${1}Corpus${2}')
      $x = [regex]::Replace($x, '(<Company>)[^<]*(</Company>)', '${1}${2}')
      $e.Delete(); $ne = $zip.CreateEntry($n); $sw = New-Object IO.StreamWriter($ne.Open(), (New-Object Text.UTF8Encoding($false))); $sw.Write($x); $sw.Close()
    }
  } finally { $zip.Dispose() }
}

$s0 = Start-PNWord 900
try {
  $w = $s0.App
  foreach ($k in $sc.Keys) {
    if ($Only -and ($Only -split ',') -notcontains $k) { continue }
    $path = Join-Path $Out "$k.docx"
    $d = $w.Documents.Add()
    try {
      $sel = $w.Selection
      & $sc[$k] $w $d $sel
      $bp = $d.BuiltInDocumentProperties; $it = [System.__ComObject].InvokeMember('Item','GetProperty',$null,$bp,@('Author')); [System.__ComObject].InvokeMember('Value','SetProperty',$null,$it,@('Corpus'))
      $d.SaveAs2([string]$path, 12)
      Write-Host "ok $k"
    } catch { Write-Host "FALLITO $k : $_" } finally { $d.Close(0) }
    if (Test-Path $path) { Patch-Props $path }
  }
} finally { Stop-PNWord $s0 }
