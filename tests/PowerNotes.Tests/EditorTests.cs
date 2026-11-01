using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Fonts;
using PowerNotes.Core.Layout;
using PowerNotes.Core.Model;
using PowerNotes.Core.Pdf;

namespace PowerNotes.Tests;

public class EditorTests
{
    static (WordDoc d, EditorSession s) NewDoc() { var d = WordDoc.CreateNew(); return (d, new EditorSession(d)); }
    static string AllText(WordDoc d) => string.Join("\n", d.Paragraphs().Select(ParaText.Text));

    static byte[] TinyPng()
    {
        // PNG 2x2 valido
        return Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4nGP8z8Dwn4GBgYEJRABCAAUEAQHBXQbyAAAAAElFTkSuQmCC");
    }

    [Fact]
    public void Font_Measure_MatchesRealMetrics()
    {
        // Arial: H=1479, e=1139, l=455, l=455, o=1139 su 2048 → 2.278 em
        Assert.Equal(22.78, FontCatalog.Measure("Hello", "Arial", 10, false, false), 1);
        Assert.True(FontCatalog.Measure("Hello", "Arial", 10, true, false) > FontCatalog.Measure("Hello", "Arial", 10, false, false));
        Assert.Equal(11.5, FontCatalog.LineHeight("Arial", 10, false, false), 1);   // hhea ascent+descent+lineGap
    }

    [Fact]
    public void Typing_Enter_Backspace_UndoRedo()
    {
        var (d, s) = NewDoc();
        s.Type("Ciao mondo");
        s.Enter();
        s.Type("Seconda");
        Assert.Equal("Ciao mondo\nSeconda", AllText(d));
        s.MoveHorizontal(-1, false, true);               // parola precedente
        Assert.Equal(0, s.Caret.Off);
        s.Backspace();                                   // unisce con il paragrafo precedente
        Assert.Equal("Ciao mondoSeconda", AllText(d));
        Assert.True(s.Undo());
        Assert.Equal("Ciao mondo\nSeconda", AllText(d));
        Assert.True(s.Redo());
        Assert.Equal("Ciao mondoSeconda", AllText(d));
    }

    [Fact]
    public void Typing_IsCoalescedInOneUndoStep()
    {
        var (d, s) = NewDoc();
        foreach (var c in "abcdef") s.Type(c.ToString());
        Assert.Equal("abcdef", AllText(d));
        s.Undo();
        Assert.Equal("", AllText(d));
    }

    [Fact]
    public void DeleteAcrossParagraphs_MergesAndKeepsRest()
    {
        var (d, s) = NewDoc();
        s.Type("uno\ndue\ntre");
        var ps = s.Paragraphs();
        s.SetCaret(ps[0], 1); s.SetCaret(ps[2], 2, true);
        s.DeleteSelection();
        Assert.Equal("ue", AllText(d));
    }

    [Fact]
    public void CharacterFormatting_BoldItalicColorSizeFont_Highlight()
    {
        var (d, s) = NewDoc();
        s.Type("normale grassetto fine");
        var p = s.Paragraphs()[0];
        s.SetCaret(p, 9); s.SetCaret(p, 18, true);
        s.ToggleBold(); s.ToggleItalic(); s.ToggleUnderline();
        s.SetColor("FF0000"); s.SetSize(20); s.SetFont("Times New Roman"); s.SetHighlight(HighlightColorValues.Yellow);
        Assert.Equal("normale grassetto fine", AllText(d));
        s.SetCaret(p, 12);
        var f = s.FormatAtCaret();
        Assert.True(f.Bold && f.Italic && f.Underline);
        Assert.Equal("FF0000", f.Color); Assert.Equal(20, f.Size); Assert.Equal("Times New Roman", f.Font); Assert.Equal("yellow", f.Highlight);
        s.SetCaret(p, 3);
        Assert.False(s.FormatAtCaret().Bold);
        s.SetCaret(p, 9); s.SetCaret(p, 18, true);
        s.ToggleBold();                                  // seconda volta: lo toglie
        s.SetCaret(p, 12); Assert.False(s.FormatAtCaret().Bold);
    }

    [Fact]
    public void PendingFormat_AppliesToNextTypedText()
    {
        var (d, s) = NewDoc();
        s.Type("a");
        s.ToggleBold();
        s.Type("b");
        s.ToggleBold();
        s.Type("c");
        var p = s.Paragraphs()[0];
        var runs = p.Elements<Run>().ToList();
        Assert.Equal("abc", ParaText.Text(p));
        s.SetCaret(p, 2); Assert.True(s.FormatAtCaret().Bold);      // dopo 'b'
        s.SetCaret(p, 1); Assert.False(s.FormatAtCaret().Bold);     // dopo 'a'
        s.SetCaret(p, 3); Assert.False(s.FormatAtCaret().Bold);
    }

    [Fact]
    public void ParagraphFormatting_AlignmentIndentSpacing_ReflectedInLayout()
    {
        var (d, s) = NewDoc();
        s.Type("Testo di prova per il paragrafo");
        s.SetAlignment(Align.Center);
        s.SetIndent(36, 18, -18);
        s.SetSpacing(6, 12); s.SetLineSpacing(1.5);
        var pf = s.ParaFormatAtCaret();
        Assert.Equal(Align.Center, pf.Alignment); Assert.Equal(36, pf.IndLeft); Assert.Equal(18, pf.IndRight); Assert.Equal(-18, pf.FirstLine);
        Assert.Equal(6, pf.Before); Assert.Equal(12, pf.After); Assert.Equal(360, pf.Line);
        var line = s.Layout.LinesOf(s.Paragraphs()[0])[0];
        Assert.True(line.Segs[0].X > 72 + 18 - 1);       // spostata da rientro e centratura
    }

    [Fact]
    public void Styles_Heading1_ChangesFontAndSize()
    {
        var (d, s) = NewDoc();
        s.Type("Titolo");
        s.SetStyle("Heading1");
        var f = s.FormatAtCaret();
        Assert.Equal(16, f.Size);
        var pf = s.ParaFormatAtCaret();
        Assert.True(pf.KeepNext);
        s.SetStyle("Normal");
        Assert.Equal(11, s.FormatAtCaret().Size);
    }

    [Fact]
    public void Lists_BulletAndNumbered_ProduceLabels()
    {
        var (d, s) = NewDoc();
        s.Type("uno\ndue\ntre");
        s.SelectAll();
        s.ToggleList("decimal");
        var lines = s.Layout.Pages[0].Lines;
        var labels = lines.Select(l => l.Segs.FirstOrDefault(x => x.Synthetic)?.Text).ToList();
        Assert.Equal(new[] { "1.", "2.", "3." }, labels);
        s.ToggleList("decimal");                         // di nuovo: rimuove
        Assert.All(s.Layout.Pages[0].Lines, l => Assert.DoesNotContain(l.Segs, x => x.Synthetic));
        s.ToggleList("bullet");
        Assert.All(s.Layout.Pages[0].Lines, l => Assert.Equal("•", l.Segs.First(x => x.Synthetic).Text));
    }

    [Fact]
    public void Lists_MultilevelNumbering()
    {
        var (d, s) = NewDoc();
        s.Type("a\nb\nc\nd");
        s.SelectAll(); s.ToggleList("decimal");
        var ps = s.Paragraphs();
        s.SetCaret(ps[1], 0); s.ChangeIndent(36);        // b → livello 2
        s.SetCaret(ps[2], 0); s.ChangeIndent(36);
        var labels = s.Layout.Pages[0].Lines.Select(l => l.Segs.First(x => x.Synthetic).Text).ToList();
        Assert.Equal(new[] { "1.", "a.", "b.", "2." }, labels);
    }

    [Fact]
    public void Table_Insert_Merge_AddRow_Navigate()
    {
        var (d, s) = NewDoc();
        s.Type("prima");
        s.Enter();
        s.InsertTable(3, 3);
        Assert.Single(d.Body.Elements<Table>());
        Assert.True(s.InTable);
        var cells = d.Body.Descendants<TableCell>().ToList();
        Assert.Equal(9, cells.Count);
        // seleziona la prima riga e unisci
        var first = cells[0].Descendants<Paragraph>().First(); var third = cells[2].Descendants<Paragraph>().First();
        s.SetCaret(first, 0); s.SetCaret(third, 0, true);
        Assert.True(s.MergeCells());
        var row0 = d.Body.Descendants<TableRow>().First().Elements<TableCell>().ToList();
        Assert.Single(row0);
        Assert.Equal(3, row0[0].TableCellProperties!.GridSpan!.Val!.Value);
        // layout: la cella unita è larga come la tabella
        var lay = d.Layout();
        Assert.True(lay.Pages[0].Prims.OfType<LinePrim>().Any());
        Assert.True(s.AddRowBelow());
        Assert.Equal(4, d.Body.Descendants<TableRow>().Count());
        Assert.True(s.AddColumnRight());
        Assert.Equal(4, d.Body.Descendants<TableGrid>().First().Elements<GridColumn>().Count());
        s.TabInTable(false);
        Assert.NotSame(first, s.Caret.P);
        // salvato è valido
        var path = Path.Combine(Path.GetTempPath(), "pn-table-" + Guid.NewGuid().ToString("N") + ".docx");
        try { d.Save(path); using var pk = WordprocessingDocument.Open(path, false); Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(pk)); }
        finally { File.Delete(path); }
    }

    [Fact]
    public void VerticalMerge_SpansRows()
    {
        var (d, s) = NewDoc();
        s.InsertTable(3, 2);
        var cells = d.Body.Descendants<TableCell>().ToList();
        s.SetCaret(cells[0].Descendants<Paragraph>().First(), 0); s.SetCaret(cells[2].Descendants<Paragraph>().First(), 0, true);
        Assert.True(s.MergeCells());
        var rows = d.Body.Descendants<TableRow>().ToList();
        Assert.Equal(MergedCellValues.Restart, rows[0].Elements<TableCell>().First().TableCellProperties!.VerticalMerge!.Val!.Value);
        Assert.NotNull(rows[1].Elements<TableCell>().First().TableCellProperties!.VerticalMerge);
        Assert.NotEmpty(d.Layout().Pages);
    }

    [Fact]
    public void Image_Insert_RendersInlineAndSurvivesSave()
    {
        var (d, s) = NewDoc();
        var png = Path.Combine(Path.GetTempPath(), "pn-" + Guid.NewGuid().ToString("N") + ".png");
        File.WriteAllBytes(png, TinyPng());
        s.Type("x");
        s.InsertImage(png);
        var lay = d.Layout();
        var img = lay.Pages[0].Prims.OfType<ImagePrim>().Single();
        Assert.NotNull(img.Part);
        Assert.Equal(1.5, img.W, 1);                      // 2px a 96 dpi = 1,5 pt
        var path = Path.Combine(Path.GetTempPath(), "pn-img-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            d.Save(path);
            using var re = WordDoc.Open(path);
            Assert.Single(re.Layout().Pages[0].Prims.OfType<ImagePrim>());
            using var pk = WordprocessingDocument.Open(path, false);
            Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(pk));
        }
        finally { File.Delete(png); File.Delete(path); }
        Assert.Throws<InvalidOperationException>(() => { var f = Path.GetTempFileName(); try { File.WriteAllText(f, "non immagine"); s.InsertImage(f); } finally { File.Delete(f); } });
    }

    [Fact]
    public void PageSetup_LandscapeMarginsAndSize()
    {
        var (d, s) = NewDoc();
        s.Type("x");
        var p = s.GetPageSetup();
        s.SetPageSetup(p with { Landscape = true, Left = 36, Right = 36 });
        var pg = d.Layout().Pages[0];
        Assert.True(pg.Width > pg.Height);
        var line = d.Layout().LinesOf(s.Paragraphs()[0])[0];
        Assert.Equal(36, line.Left, 1);
        Assert.Throws<ArgumentException>(() => s.SetPageSetup(p with { Left = 500, Right = 500 }));
    }

    [Fact]
    public void PageBreak_And_Pagination_SplitAcrossPages()
    {
        var (d, s) = NewDoc();
        s.Type("pagina uno");
        s.InsertPageBreak();
        s.Type("pagina due");
        var lay = d.Layout();
        Assert.Equal(2, lay.Pages.Count);
        Assert.Contains(lay.Pages[1].Lines, l => l.Segs.Any(x => x.Text.Contains("due")));
        var big = string.Join("\n", Enumerable.Range(0, 120).Select(i => "riga " + i));
        s.Type("\n" + big);
        Assert.True(d.Layout().Pages.Count >= 3);
        // nessuna riga oltre il margine inferiore
        foreach (var pg in d.Layout().Pages) Assert.All(pg.Lines, l => Assert.True(l.Y + l.Height <= pg.Height - 0.1));
    }

    [Fact]
    public void HeaderFooter_PageNumbers_AreComputedPerPage()
    {
        var (d, s) = NewDoc();
        s.Type("uno"); s.InsertPageBreak(); s.Type("due"); s.InsertPageBreak(); s.Type("tre");
        s.SetHeaderFooter(false, "", Align.Center, PageNumFormat.PageOfTotal);
        s.SetHeaderFooter(true, "Intestazione", Align.Left, PageNumFormat.None);
        var lay = d.Layout();
        Assert.Equal(3, lay.Pages.Count);
        for (int i = 0; i < 3; i++)
        {
            var txt = string.Concat(lay.Pages[i].Prims.OfType<TextPrim>().Select(t => t.Text + " "));
            Assert.Contains("Intestazione", txt);
            Assert.Contains("Pagina", txt);
            Assert.Contains((i + 1).ToString(), txt);
            Assert.Contains("3", txt);
        }
        Assert.Equal("Intestazione", s.GetHeaderFooterText(true));
        var path = Path.Combine(Path.GetTempPath(), "pn-hf-" + Guid.NewGuid().ToString("N") + ".docx");
        try { d.Save(path); using var pk = WordprocessingDocument.Open(path, false); Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(pk)); }
        finally { File.Delete(path); }
    }

    [Fact]
    public void FindReplace()
    {
        var (d, s) = NewDoc();
        s.Type("Gatto gatto GATTO e cane");
        Assert.True(s.FindNext("gatto", false)); Assert.Equal(5, s.Caret.Off - s.Anchor.Off + 0);
        Assert.False(s.FindNext("zebra", false));
        Assert.Equal(1, s.ReplaceAll("gatto", "cane", true));
        Assert.Equal("Gatto cane GATTO e cane", AllText(d));
        Assert.Equal(2, s.ReplaceAll("gatto", "x", false));
        Assert.Equal("x cane x e cane", AllText(d));
        s.Undo(); Assert.Equal("Gatto cane GATTO e cane", AllText(d));
    }

    [Fact]
    public void Justify_FillsLineWidth()
    {
        var (d, s) = NewDoc();
        s.Type(string.Join(" ", Enumerable.Repeat("parola", 40)));
        s.SetAlignment(Align.Justify);
        var lines = s.Layout.LinesOf(s.Paragraphs()[0]);
        Assert.True(lines.Count > 2);
        var first = lines[0];
        var lastSeg = first.Segs.Last(x => x.Text.Trim().Length > 0);
        Assert.Equal(first.Right, lastSeg.X + lastSeg.Width, 1);
        Assert.True(lines[^1].Segs.Last(x => x.Text.Trim().Length > 0).X + 5 < lines[^1].Right);   // l'ultima riga no
    }

    [Fact]
    public void HitTest_RoundTripsCaretPositions()
    {
        var (d, s) = NewDoc();
        s.Type("Prima riga di testo abbastanza lunga da andare a capo almeno una volta in questa pagina di prova, davvero.");
        s.Enter(); s.Type("Seconda");
        var L = s.Layout;
        foreach (var p in s.Paragraphs())
            for (int off = 0; off <= ParaText.Length(p); off += 7)
            {
                var line = L.Locate(p, off)!;
                double x = line.XAt(off) + 0.01, y = line.Y + line.Height / 2;
                var hit = L.HitTest(line.PageIndex, x, y)!.Value;
                Assert.Equal(p, hit.Para); Assert.Equal(off, hit.Offset);
            }
    }

    [Fact]
    public void Navigation_UpDownHomeEnd()
    {
        var (d, s) = NewDoc();
        s.Type("xxx\nxxx\nxxx");
        var ps = s.Paragraphs();
        s.SetCaret(ps[2], 2);
        s.MoveVertical(-1, false); Assert.Same(ps[1], s.Caret.P); Assert.Equal(2, s.Caret.Off);
        s.MoveVertical(-1, false); Assert.Same(ps[0], s.Caret.P);
        s.MoveVertical(1, false); s.MoveVertical(1, false); Assert.Same(ps[2], s.Caret.P);
        s.MoveLineEdge(false, false); Assert.Equal(0, s.Caret.Off);
        s.MoveLineEdge(true, false); Assert.Equal(3, s.Caret.Off);
        s.MoveLineEdge(false, true, doc: true); Assert.Equal(ps[0], s.Caret.P); Assert.True(s.HasSelection);
    }

    [Fact]
    public void Pdf_Export_HasPagesAndText()
    {
        var (d, s) = NewDoc();
        s.Type("Documento PDF"); s.InsertPageBreak(); s.Type("Seconda pagina");
        var bytes = PdfExporter.Export(d.Layout(), "t");
        Assert.StartsWith("%PDF", System.Text.Encoding.ASCII.GetString(bytes, 0, 5));
        var txt = System.Text.Encoding.Latin1.GetString(bytes);
        Assert.Contains("/Type /Pages", txt.Replace("/Type/Pages", "/Type /Pages"));
        Assert.Matches(@"/Count 2\b", txt);
    }

    [Fact]
    public void NewDocument_IsValidOpenXml_AndReopens()
    {
        var (d, s) = NewDoc();
        s.Type("Contenuto");
        var path = Path.Combine(Path.GetTempPath(), "pn-new-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            d.Save(path);
            using (var pk = WordprocessingDocument.Open(path, false)) Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(pk));
            using var re = WordDoc.Open(path); Assert.Equal("Contenuto", AllText(re));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void PdfExport_OfEveryCorpusFile_DoesNotThrow()
    {
        foreach (var f in TestPaths.CorpusFiles())
        {
            using var d = WordDoc.Open(f);
            var bytes = PdfExporter.Export(d.Layout(), Path.GetFileName(f));
            Assert.True(bytes.Length > 500, f);
        }
    }

    [Fact]
    public void StrikeSuperSubscript_And_CutPaste()
    {
        var (d, s) = NewDoc();
        s.Type("abcdef");
        var p = s.Paragraphs()[0];
        s.SetCaret(p, 1); s.SetCaret(p, 3, true); s.ToggleStrike();
        s.SetCaret(p, 2); Assert.True(s.FormatAtCaret().Strike);
        s.SetCaret(p, 3); s.SetCaret(p, 5, true); s.SetVertAlign(1);
        s.SetCaret(p, 5); Assert.Equal(1, s.FormatAtCaret().VertAlign);
        s.SetCaret(p, 3); s.SetCaret(p, 5, true); s.SetVertAlign(-1);
        s.SetCaret(p, 5); Assert.Equal(-1, s.FormatAtCaret().VertAlign);
        s.SetCaret(p, 1); s.SetCaret(p, 3, true);
        var cut = s.Cut();
        Assert.Equal("bc", cut); Assert.Equal("adef", AllText(d));
        s.SetCaret(p, 4); s.Type(cut);
        Assert.Equal("adefbc", AllText(d));
    }

    [Fact]
    public void SpaceBetweenParagraphs_UsesLargerOfAfterAndBefore()
    {
        var (d, s) = NewDoc();
        s.Type("a" + (char)10 + "b");
        var ps = s.Paragraphs();
        s.SetCaret(ps[0], 0); s.SetSpacing(0, 20);
        s.SetCaret(ps[1], 0); s.SetSpacing(12, 0);
        var L = d.Layout();
        double gap = L.LinesOf(ps[1])[0].Y - (L.LinesOf(ps[0])[0].Y + L.LinesOf(ps[0])[0].Height);
        Assert.Equal(20, gap, 1);      // max(20, 12), come Word
    }
}

public class SectionColumnTests
{
    [Fact]
    public void Columns_And_SectionBreaks_ValidAndLaidOut()
    {
        var d = WordDoc.CreateNew(); var s = new EditorSession(d);
        s.Type(string.Join("\n", Enumerable.Repeat("riga di testo piuttosto lunga per riempire le colonne della pagina di prova", 60)));
        s.SetColumns(2);
        var lay = d.Layout();
        var xs = lay.Pages[0].Lines.Select(l => Math.Round(l.Left)).Distinct().ToList();
        Assert.Equal(2, xs.Count);
        s.SetColumns(1);
        s.SetCaret(s.Paragraphs()[3], 5);
        s.InsertSectionBreak(false);
        Assert.Equal(2, d.Body.Descendants<SectionProperties>().Count());
        Assert.True(d.Layout().Pages.Count >= 2);
        s.InsertSectionBreak(true);
        Assert.Equal(3, d.Body.Descendants<SectionProperties>().Count());
        var path = Path.Combine(Path.GetTempPath(), "pn-sec-" + Guid.NewGuid().ToString("N") + ".docx");
        try { d.Save(path); using var pk = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, false); Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(pk)); }
        finally { File.Delete(path); }
    }
}

public class ClipboardTests
{
    [Fact]
    public void Fragment_KeepsFormattingAcrossParagraphsAndDocuments()
    {
        var d = WordDoc.CreateNew(); var s = new EditorSession(d);
        s.Type("uno due\ntre quattro");
        var ps = s.Paragraphs();
        s.SetCaret(ps[0], 4); s.SetCaret(ps[1], 3, true);
        s.ToggleBold(); s.SetColor("FF0000");
        s.SetCaret(ps[0], 4); s.SetCaret(ps[1], 3, true);
        var frag = s.CopyFragment()!;
        Assert.Contains("FF0000", frag);
        var d2 = WordDoc.CreateNew(); var s2 = new EditorSession(d2);
        s2.Type("AB");
        s2.SetCaret(s2.Paragraphs()[0], 1);
        s2.PasteFragment(frag, false);
        Assert.Equal("Adue\ntreB", string.Join("\n", d2.Paragraphs().Select(ParaText.Text)));
        var p = d2.Paragraphs()[0]; s2.SetCaret(p, 3);
        Assert.True(s2.FormatAtCaret().Bold); Assert.Equal("FF0000", s2.FormatAtCaret().Color);
        s2.Undo();
        Assert.Equal("AB", string.Join("\n", d2.Paragraphs().Select(ParaText.Text)));
    }
}
