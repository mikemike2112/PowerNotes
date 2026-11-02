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
}
