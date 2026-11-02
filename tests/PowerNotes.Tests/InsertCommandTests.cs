using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Layout;
using PowerNotes.Core.Model;

namespace PowerNotes.Tests;

/// <summary>Collegamento, Segnalibro, Nota di chiusura, Nota successiva.</summary>
public class InsertCommandTests
{
    static (WordDoc d, EditorSession s) NewDoc() => TableCommandTests.NewDoc();
    static int Snap(WordDoc d) => d.Body.OuterXml.GetHashCode();
    static string AllText(WordDoc d) => string.Join("\n", d.Paragraphs().Select(ParaText.Text));

    static WordDoc Reopen(WordDoc d)
    {
        var path = Path.Combine(Path.GetTempPath(), "pn-ins-" + Guid.NewGuid().ToString("N") + ".docx");
        d.Save(path);
        try { return WordDoc.Open(path); } finally { File.Delete(path); }
    }

    [Fact]
    public void Hyperlink_WrapsSelection_OneUndo_RoundTrips()
    {
        var (d, s) = NewDoc();
        s.Type("vai al sito ora");
        var p = s.Paragraphs()[0];
        int before = Snap(d);
        s.SetCaret(p, 7); s.SetCaret(p, 11, true);
        s.SetHyperlink("example.com/pagina");
        var h = d.Body.Descendants<Hyperlink>().Single();
        Assert.Equal("sito", string.Join("", h.Descendants<Text>().Select(t => t.Text)));
        Assert.Equal("vai al sito ora", AllText(d));
        s.SetCaret(p, 9);
        Assert.Equal("https://example.com/pagina", s.HyperlinkAt());
        var f = s.FormatAtCaret(); Assert.True(f.Underline); Assert.Equal("0563C1", f.Color);
        TableCommandTests.AssertValid(d);
        using (var d2 = Reopen(d)) { var s2 = new EditorSession(d2); var p2 = s2.Paragraphs()[0]; s2.SetCaret(p2, 9); Assert.Equal("https://example.com/pagina", s2.HyperlinkAt()); }
        Assert.True(s.Undo());
        Assert.Equal(before, Snap(d));
    }

    [Fact]
    public void Hyperlink_ChangeAndRemove()
    {
        var (d, s) = NewDoc();
        s.Type("abc def");
        var p = s.Paragraphs()[0];
        s.SetCaret(p, 0); s.SetCaret(p, 3, true); s.SetHyperlink("https://a.example/");
        s.SetCaret(p, 1);
        s.SetHyperlink("mailto:x@example.org");
        Assert.Equal("mailto:x@example.org", s.HyperlinkAt());
        Assert.Single(d.Pkg.MainDocumentPart!.HyperlinkRelationships);   // la vecchia relazione e' stata tolta
        Assert.True(s.RemoveHyperlink());
        Assert.Empty(d.Body.Descendants<Hyperlink>()); Assert.Empty(d.Pkg.MainDocumentPart.HyperlinkRelationships);
        Assert.False(s.FormatAtCaret().Underline);
        Assert.Equal("abc def", AllText(d));
        Assert.False(s.RemoveHyperlink());
        TableCommandTests.AssertValid(d);
    }

    [Fact]
    public void Hyperlink_NoSelection_InsertsTextAndToBookmark()
    {
        var (d, s) = NewDoc();
        s.Type("x ");
        s.SetHyperlink("https://example.org", "Esempio");
        Assert.Equal("x Esempio", AllText(d));
        s.Type(" fine");
        var p = s.Paragraphs()[0];
        s.SetCaret(p, 0); s.AddBookmark("inizio");
        s.SetCaret(p, 3); s.SetCaret(p, 5, true);
        s.SetHyperlink("#inizio");
        s.SetCaret(p, 4);
        Assert.Equal("#inizio", s.HyperlinkAt());
        Assert.Throws<ArgumentException>(() => s.SetHyperlink("#nonesiste"));
        TableCommandTests.AssertValid(d);
    }

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("file:///C:/Windows/system.ini")]
    [InlineData("")]
    [InlineData("http://")]
    [InlineData("ftp:")]
    [InlineData("data:text/html;base64,AAAA")]
    public void Hyperlink_UnsafeOrInvalid_Refused(string url)
    {
        var (d, s) = NewDoc();
        s.Type("testo"); var p = s.Paragraphs()[0]; s.SetCaret(p, 0); s.SetCaret(p, 5, true);
        Assert.Throws<ArgumentException>(() => s.SetHyperlink(url));
        Assert.Empty(d.Body.Descendants<Hyperlink>());
    }

    [Fact]
    public void Bookmark_AddGoToDeleteMoveAndValidate()
    {
        var (d, s) = NewDoc();
        s.Type("uno due\ntre quattro");
        var ps = s.Paragraphs();
        int before = Snap(d);
        s.SetCaret(ps[0], 4); s.SetCaret(ps[0], 7, true);
        s.AddBookmark("Parte_1");
        Assert.Equal(["Parte_1"], s.Bookmarks());
        Assert.Equal("uno due\ntre quattro", AllText(d));
        TableCommandTests.AssertValid(d);
        s.SetCaret(ps[1], 0);
        Assert.True(s.GoToBookmark("Parte_1")); Assert.Equal(ps[0], s.Caret.P); Assert.Equal(4, s.Caret.Off);
        s.SetCaret(ps[1], 4); s.AddBookmark("Parte_1");              // stesso nome: si sposta
        Assert.Single(d.Body.Descendants<BookmarkStart>());
        Assert.True(s.GoToBookmark("Parte_1")); Assert.Equal(ps[1], s.Caret.P); Assert.Equal(4, s.Caret.Off);
        using (var d2 = Reopen(d)) Assert.Equal(["Parte_1"], new EditorSession(d2).Bookmarks());
        Assert.True(s.DeleteBookmark("Parte_1")); Assert.False(s.DeleteBookmark("Parte_1"));
        Assert.Empty(d.Body.Descendants<BookmarkStart>()); Assert.Empty(d.Body.Descendants<BookmarkEnd>());
        Assert.False(s.GoToBookmark("Parte_1"));
        Assert.True(s.Undo()); Assert.Single(d.Body.Descendants<BookmarkStart>());
    }

    [Theory]
    [InlineData("")]
    [InlineData("1abc")]
    [InlineData("_nascosto")]
    [InlineData("con spazio")]
    [InlineData("troppo_lungo_troppo_lungo_troppo_lungo_troppo")]
    [InlineData("à")]
    public void Bookmark_InvalidNames_Refused(string name)
    {
        var (d, s) = NewDoc();
        s.Type("x");
        Assert.Throws<ArgumentException>(() => s.AddBookmark(name));
        Assert.Empty(d.Body.Descendants<BookmarkStart>());
    }

    [Fact]
    public void Endnote_InsertsRefAndNote_LaidOutAtEndWithRomanMark()
    {
        var (d, s) = NewDoc();
        s.Type("Prima frase.\nSeconda frase.");
        var ps = s.Paragraphs();
        int before = Snap(d);
        s.SetCaret(ps[0], 5); s.InsertEndnote("Prima nota");
        s.SetCaret(ps[1], 6); s.InsertEndnote("Seconda nota");
        Assert.Equal(2, d.Body.Descendants<EndnoteReference>().Count());
        var en = d.Pkg.MainDocumentPart!.EndnotesPart!.Endnotes!;
        Assert.Equal(4, en.Elements<Endnote>().Count());
        TableCommandTests.AssertValid(d);
        var lay = d.Layout();
        var texts = lay.Pages[^1].Prims.OfType<TextPrim>().Select(t => t.Text).ToList();
        Assert.Contains("Prima", texts); Assert.Contains("Seconda", texts); Assert.Contains("nota", texts);
        Assert.Contains("i", texts); Assert.Contains("ii", texts);
        double yBody = lay.Pages[0].Prims.OfType<TextPrim>().First(t => t.Text == "Prima").Baseline;
        double yNote = lay.Pages[0].Prims.OfType<TextPrim>().Last(t => t.Text == "nota").Baseline;
        Assert.True(yNote > yBody + 20);
        using (var d2 = Reopen(d)) Assert.Equal(2, d2.Body.Descendants<EndnoteReference>().Count());
        Assert.True(s.Undo()); Assert.True(s.Undo());
        Assert.Equal(before, Snap(d));
        Assert.Empty(d.Body.Descendants<EndnoteReference>());
    }

    [Fact]
    public void NextNote_CyclesThroughFootnotesAndEndnotes()
    {
        var (d, s) = NewDoc();
        Assert.False(s.NextNote());
        s.Type("uno due tre");
        var p = s.Paragraphs()[0];
        s.SetCaret(p, 3); s.InsertFootnote("a piè");
        s.SetCaret(p, 7); s.InsertEndnote("di chiusura");
        s.SetCaret(p, 0);
        Assert.True(s.NextNote()); Assert.Equal(3, s.Caret.Off);
        Assert.True(s.NextNote()); Assert.Equal(7, s.Caret.Off);
        Assert.True(s.NextNote()); Assert.Equal(3, s.Caret.Off);   // ricomincia dall'inizio
    }

    [Fact]
    public void Shape_InsertedAnchored_DrawnOnPageAndInPdf_OneUndo_Valid()
    {
        var (d, s) = NewDoc();
        s.Type("Testo sotto la forma");
        int before = Snap(d);
        s.InsertShape("ellipse", 100, 60, "4472C4", "2F528F");
        var shapes = d.Layout().Pages[0].Prims.OfType<ShapePrim>().ToList();
        Assert.Single(shapes);
        Assert.Equal("ellipse", shapes[0].Geom); Assert.Equal(100, shapes[0].W, 1); Assert.Equal(60, shapes[0].H, 1); Assert.Equal("4472C4", shapes[0].Fill);
        double textY = d.Layout().Pages[0].Prims.OfType<TextPrim>().First().Baseline;
        Assert.True(textY > shapes[0].Y + 60, "il testo sta sotto la forma");
        Assert.Equal("Testo sotto la forma", AllText(d));
        TableCommandTests.AssertValid(d);
        using (var d2 = Reopen(d)) Assert.Single(d2.Layout().Pages[0].Prims.OfType<ShapePrim>());
        Assert.True(PowerNotes.Core.Pdf.PdfExporter.Export(d.Layout()).Length > 500);
        Assert.True(s.Undo());
        Assert.Equal(before, Snap(d));
        Assert.Empty(d.Layout().Pages[0].Prims.OfType<ShapePrim>());
    }

    [Fact]
    public void Shape_AllGalleryKindsValidAndPolygonsClosed()
    {
        var (d, s) = NewDoc();
        foreach (var (prst, _) in Shapes.Gallery) { s.InsertShape(prst, 50, 50, "FF0000", null); s.Enter(); }
        Assert.Equal(Shapes.Gallery.Length, d.Layout().Pages.SelectMany(p => p.Prims.OfType<ShapePrim>()).Count());
        TableCommandTests.AssertValid(d);
        Assert.Equal(10, Shapes.Polygon("star5", 50, 50)!.Length);
    }

    [Theory]
    [InlineData("nonesiste", 10, 10, "FF0000")]
    [InlineData("rect", 0, 10, "FF0000")]
    [InlineData("rect", 10, 5000, "FF0000")]
    [InlineData("rect", double.NaN, 10, "FF0000")]
    [InlineData("rect", 10, 10, "ZZZZZZ")]
    public void Shape_InvalidInput_Refused(string prst, double w, double h, string fill)
    {
        var (d, s) = NewDoc();
        Assert.ThrowsAny<ArgumentException>(() => s.InsertShape(prst, w, h, fill, null));
        Assert.Empty(d.Body.Descendants<Drawing>());
    }

    [Fact]
    public void ImageData_InsertsWithRequestedWidth()
    {
        var (d, s) = NewDoc();
        var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4nGP8z8Dwn4GBgYEJRABCAAUEAQHBXQbyAAAAAElFTkSuQmCC");
        s.InsertImageData(png, "Icona: prova", 40);
        var ext = d.Body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().Single();
        Assert.Equal(40 * 12700, ext.Cx!.Value);
        TableCommandTests.AssertValid(d);
    }

    [Fact]
    public void LineNumbers_DrawnInMarginPerLine_RestartModes_OneUndo_Valid()
    {
        var (d, s) = NewDoc();
        s.Type(string.Join("\n", Enumerable.Range(1, 5).Select(i => "riga " + i)));
        int before = Snap(d);
        Assert.Equal(EditorSession.LineNumMode.None, s.GetLineNumbers());
        s.SetLineNumbers(EditorSession.LineNumMode.Continuous);
        Assert.Equal(EditorSession.LineNumMode.Continuous, s.GetLineNumbers());
        var nums = d.Layout().Pages[0].Prims.OfType<TextPrim>().Where(t => t.Text.All(char.IsDigit) && t.X < 60).ToList();
        Assert.Equal(["1", "2", "3", "4", "5"], nums.Select(t => t.Text));
        Assert.All(nums, t => Assert.True(t.X + t.Width < d.Layout().Pages[0].MarL));
        TableCommandTests.AssertValid(d);
        using (var d2 = Reopen(d)) Assert.Equal(5, d2.Layout().Pages[0].Prims.OfType<TextPrim>().Count(t => t.Text.All(char.IsDigit) && t.X < 60));
        // un paragrafo che sopprime i numeri non viene contato
        var ps = s.Paragraphs();
        XmlOps.PPr(ps[2]).SuppressLineNumbers = new SuppressLineNumbers();
        d.Touched();
        Assert.Equal(4, d.Layout().Pages[0].Prims.OfType<TextPrim>().Count(t => t.Text.All(char.IsDigit) && t.X < 60));
        s.SetLineNumbers(EditorSession.LineNumMode.None);
        Assert.Empty(d.Layout().Pages[0].Prims.OfType<TextPrim>().Where(t => t.Text.All(char.IsDigit) && t.X < 60));
        Assert.Empty(d.Body.Descendants<LineNumberType>());
        s.SetLineNumbers(EditorSession.LineNumMode.RestartPage);
        Assert.True(s.Undo()); Assert.True(s.Undo());
    }

    [Fact]
    public void LineNumbers_RestartPageVsContinuous_AcrossPages()
    {
        var (d, s) = NewDoc();
        s.Type(string.Join("\n", Enumerable.Range(1, 100).Select(i => "riga " + i)));
        s.SetLineNumbers(EditorSession.LineNumMode.RestartPage);
        var lay = d.Layout();
        Assert.True(lay.Pages.Count >= 2);
        Assert.Equal("1", lay.Pages[1].Prims.OfType<TextPrim>().First(t => t.Text.All(char.IsDigit) && t.X < lay.Pages[1].MarL).Text);
        s.SetLineNumbers(EditorSession.LineNumMode.Continuous);
        var first2 = d.Layout().Pages[1].Prims.OfType<TextPrim>().First(t => t.Text.All(char.IsDigit) && t.X < d.Layout().Pages[1].MarL).Text;
        Assert.NotEqual("1", first2);
    }

    [Fact]
    public void Headings_ListsOutlineParagraphsWithLevels()
    {
        var (d, s) = NewDoc();
        s.Type("Capitolo"); s.SetStyle("Heading1"); s.Enter();
        s.Type("corpo"); s.SetStyle("Normal"); s.Enter();
        s.Type("Sezione"); s.SetStyle("Heading2"); s.Enter(); s.Type("fine");
        var h = s.Headings();
        Assert.Equal([("Capitolo", 0), ("Sezione", 1)], h.Select(x => (x.Text, x.Level)).ToArray());
    }
}
