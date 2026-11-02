using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Layout;
using PowerNotes.Core.Model;

namespace PowerNotes.Tests;

/// <summary>Comandi della scheda Home che erano inerti: Copia formato, Aa, elenco a più livelli, Ordina, Sfondo, Bordi.</summary>
public class HomeCommandTests
{
    static (WordDoc d, EditorSession s) NewDoc() => TableCommandTests.NewDoc();
    static int Snap(WordDoc d) => d.Body.OuterXml.GetHashCode();
    static string AllText(WordDoc d) => string.Join("\n", d.Paragraphs().Select(ParaText.Text));

    [Fact]
    public void CopyFormat_TransfersDirectFormatting_OneUndo()
    {
        var (d, s) = NewDoc();
        s.Type("Rosso grassetto e testo normale");
        var p = s.Paragraphs()[0];
        s.SetCaret(p, 0); s.SetCaret(p, 6, true);
        s.ToggleBold(); s.SetColor("C00000"); s.SetSize(18);
        s.SetCaret(p, 3);
        var fmt = s.CaptureFormat();
        s.SetCaret(p, 22); s.SetCaret(p, 27, true);
        int before = Snap(d);
        s.ApplyFormat(fmt);
        s.SetCaret(p, 24);
        var f = s.FormatAtCaret();
        Assert.True(f.Bold); Assert.Equal("C00000", f.Color); Assert.Equal(18, f.Size);
        s.SetCaret(p, 10); Assert.False(s.FormatAtCaret().Bold);       // il resto non cambia
        Assert.Equal("Rosso grassetto e testo normale", AllText(d));
        TableCommandTests.AssertValid(d);
        Assert.True(s.Undo()); Assert.Equal(before, Snap(d));
    }

    [Theory]
    [InlineData(CaseMode.Upper, "CIAO MONDO. COME VA?")]
    [InlineData(CaseMode.Lower, "ciao mondo. come va?")]
    [InlineData(CaseMode.Title, "Ciao Mondo. Come Va?")]
    [InlineData(CaseMode.Sentence, "Ciao mondo. Come va?")]
    [InlineData(CaseMode.Toggle, "cIAO mondo. COME va?")]
    public void ChangeCase_AllModes_KeepFormattingAndLength(CaseMode mode, string expected)
    {
        var (d, s) = NewDoc();
        s.Type("Ciao MONDO. come VA?");
        var p = s.Paragraphs()[0];
        s.SetCaret(p, 5); s.SetCaret(p, 10, true); s.ToggleBold();     // «MONDO» in grassetto: più run
        s.SelectAll();
        int before = Snap(d);
        s.ChangeCase(mode);
        Assert.Equal(expected, AllText(d));
        s.SetCaret(p, 7); Assert.True(s.FormatAtCaret().Bold);
        TableCommandTests.AssertValid(d);
        Assert.True(s.Undo()); Assert.Equal(before, Snap(d));
    }

    [Fact]
    public void ChangeCase_NoSelection_ChangesWordAtCaret_AndPartialSelection()
    {
        var (d, s) = NewDoc();
        s.Type("uno due tre");
        var p = s.Paragraphs()[0];
        s.SetCaret(p, 5);
        s.ChangeCase(CaseMode.Upper);
        Assert.Equal("uno DUE tre", AllText(d));
        s.SetCaret(p, 0); s.SetCaret(p, 2, true);
        s.ChangeCase(CaseMode.Title);
        Assert.Equal("Uno DUE tre", AllText(d));
    }

    [Fact]
    public void MultilevelList_ProducesNestedLabels()
    {
        var (d, s) = NewDoc();
        s.Type("uno\ndue\ntre");
        s.SelectAll();
        s.ToggleList("multilevel");
        var ps = s.Paragraphs();
        s.SetCaret(ps[1], 0); s.ChangeIndent(36);          // secondo livello
        var labels = ps.Select(p => s.Layout.LinesOf(p)[0].Segs.First(x => x.Synthetic).Text).ToList();
        Assert.Equal(["1.", "1.1.", "2."], labels);
        TableCommandTests.AssertValid(d);
    }

    [Fact]
    public void SortParagraphs_AscDesc_OneUndo_RefusesAcrossTable()
    {
        var (d, s) = NewDoc();
        s.Type("pera\nMela\nbanana");
        s.SelectAll();
        int before = Snap(d);
        s.SortParagraphs(false);
        Assert.Equal("banana\nMela\npera", AllText(d));
        s.SortParagraphs(true);
        Assert.Equal("pera\nMela\nbanana", AllText(d));
        TableCommandTests.AssertValid(d);
        Assert.True(s.Undo()); Assert.Equal("banana\nMela\npera", AllText(d));
        var (d2, s2) = NewDoc();
        s2.Type("solo uno");
        Assert.Throws<InvalidOperationException>(() => s2.SortParagraphs(false));
    }

    [Fact]
    public void ParaShadingAndBorders_DrawnBehindText_AndRoundTrip()
    {
        var (d, s) = NewDoc();
        s.Type("Paragrafo evidenziato");
        s.SetParaShading("FFFF00");
        s.SetParaBorders(BorderEdge.Box);
        var pg = d.Layout().Pages[0];
        Assert.Contains(pg.Prims.OfType<RectPrim>(), r => r.Fill == "FFFF00");
        Assert.Equal(4, pg.Prims.OfType<LinePrim>().Count());
        int rectIdx = pg.Prims.FindIndex(x => x is RectPrim), textIdx = pg.Prims.FindIndex(x => x is TextPrim);
        Assert.True(rectIdx < textIdx);
        TableCommandTests.AssertValid(d);
        s.SetParaBorders(BorderEdge.Top);                         // con il contorno completo, un clic su un lato lo toglie
        Assert.Equal(3, d.Layout().Pages[0].Prims.OfType<LinePrim>().Count());
        s.SetParaBorders(BorderEdge.None); s.SetParaShading(null);
        Assert.DoesNotContain(d.Layout().Pages[0].Prims, x => x is LinePrim); Assert.DoesNotContain(d.Layout().Pages[0].Prims.OfType<RectPrim>(), r => r.Fill == "FFFF00");
        Assert.Null(d.Paragraphs()[0].ParagraphProperties?.ParagraphBorders);
        Assert.Throws<ArgumentException>(() => s.SetParaShading("rosso"));
        TableCommandTests.AssertValid(d);
    }

    [Fact]
    public void ParaBorders_RoundTripKeepsBordersAfterSave()
    {
        var (d, s) = NewDoc();
        s.Type("a\nb");
        s.SelectAll();
        s.SetParaBorders(BorderEdge.Bottom);
        var path = Path.Combine(Path.GetTempPath(), "pn-home-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            d.Save(path);
            using var d2 = WordDoc.Open(path);
            Assert.Equal(2, d2.Layout().Pages[0].Prims.OfType<LinePrim>().Count());
        }
        finally { File.Delete(path); }
    }
}
