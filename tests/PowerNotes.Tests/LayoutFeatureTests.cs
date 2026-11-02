using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Model;
using PowerNotes.Core.Pdf;

namespace PowerNotes.Tests;

public class LayoutFeatureTests
{
    static Paragraph P(string t) => new(new Run(new Text(t) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));

    [Fact]
    public void Footnote_IsPlacedAtBottomOfReferencePage()
    {
        using var d = WordDoc.CreateNew();
        var body = d.Body;
        body.RemoveAllChildren<Paragraph>();
        var sect = body.GetFirstChild<SectionProperties>();
        for (int i = 0; i < 120; i++)
        {
            var p = P("Riga di corpo numero " + i);
            if (i == 70) p.Append(new Run(new FootnoteReference { Id = 1 }));
            if (sect != null) body.InsertBefore(p, sect); else body.Append(p);
        }
        var fp = d.Pkg.MainDocumentPart!.AddNewPart<FootnotesPart>();
        fp.Footnotes = new Footnotes(
            new Footnote { Type = FootnoteEndnoteValues.Separator, Id = -1 },
            new Footnote(new Paragraph(new Run(new FootnoteReferenceMark()), new Run(new Text(" TestoNotaUnica") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }))) { Id = 1 });
        var res = d.Layout();
        var note = res.Pages.SelectMany(pg => pg.Lines).Single(l => ParaText.Text(l.Para).Contains("TestoNotaUnica"));
        var refLine = res.Pages.SelectMany(pg => pg.Lines).Single(l => ParaText.Text(l.Para).StartsWith("Riga di corpo numero 70"));
        Assert.Equal(refLine.PageIndex, note.PageIndex);
        var pageLines = res.Pages[note.PageIndex].Lines.Where(l => l.Story == PowerNotes.Core.Layout.Story.Body && l != note && !ReferenceEquals(l.Para, note.Para));
        Assert.All(pageLines, l => Assert.True(l.Y + l.Height <= note.Y + 0.5));
        Assert.True(note.Y + note.Height <= res.Pages[note.PageIndex].Height - res.Pages[note.PageIndex].MarB + 0.5);
        Assert.True(res.Pages[note.PageIndex].Prims.OfType<PowerNotes.Core.Layout.LinePrim>().Any(), "separatore");
        Assert.True(PdfExporter.Export(res, "n").Length > 1000);
    }

    [Fact]
    public void TableRow_TallerThanPage_SplitsAcrossPages()
    {
        using var d = WordDoc.CreateNew();
        var body = d.Body;
        body.RemoveAllChildren<Paragraph>();
        var cell = new TableCell();
        for (int i = 0; i < 150; i++) cell.Append(P("Cella riga " + i));
        var tbl = new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "5000" }), new TableRow(cell));
        var sect = body.GetFirstChild<SectionProperties>();
        if (sect != null) body.InsertBefore(tbl, sect); else body.Append(tbl);
        body.InsertBefore(P("dopo"), sect);
        var res = d.Layout();
        Assert.True(res.Pages.Count >= 3);
        foreach (var pg in res.Pages)
            Assert.All(pg.Lines, l => Assert.True(l.Y + l.Height <= pg.Height - pg.MarB + 0.5, $"riga fuori margine y={l.Y}"));
        Assert.Equal(151, res.Pages.SelectMany(pg => pg.Lines).Count(l => ParaText.Text(l.Para).StartsWith("Cella riga") || ParaText.Text(l.Para) == "dopo"));
    }
}

public class EdgeCaseTests
{
    static (WordDoc d, PowerNotes.Core.Edit.EditorSession s) NewDoc() { var d = WordDoc.CreateNew(); return (d, new PowerNotes.Core.Edit.EditorSession(d)); }
    static void ValidSave(WordDoc d)
    {
        var path = Path.Combine(Path.GetTempPath(), "pn-edge-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            d.Save(path);
            using var pk = WordprocessingDocument.Open(path, false);
            var errs = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(pk).Select(e => e.Description + " @ " + e.Path?.XPath).ToList(); Assert.True(errs.Count == 0, string.Join(" | ", errs));
            using var again = WordDoc.Open(path);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Footnote_Insert_Undo_Redo_SaveValid()
    {
        var (d, s) = NewDoc();
        s.Type("Prima del richiamo");
        s.InsertFootnote("Testo della nota");
        s.Type(" dopo");
        Assert.Single(d.Body.Descendants<FootnoteReference>());
        Assert.Equal("Prima del richiamo dopo", ParaText.Text(d.Paragraphs()[0]));
        ValidSave(d);
        var res = d.Layout();
        Assert.Contains(res.Pages[0].Lines, l => ParaText.Text(l.Para).Contains("Testo della nota"));
        s.Undo(); s.Undo();
        Assert.Empty(d.Body.Descendants<FootnoteReference>());
        s.Redo();
        Assert.Single(d.Body.Descendants<FootnoteReference>());
        s.InsertFootnote("Seconda");
        Assert.Equal(new long[] { 1, 2 }, d.Body.Descendants<FootnoteReference>().Select(r => r.Id!.Value).OrderBy(x => x).ToArray());
        ValidSave(d);
    }

    [Fact]
    public void Footnote_OnCorpusFileWithSettings_SavesValid()
    {
        using var d = WordDoc.Open(Path.Combine(TestPaths.Corpus, "plain-paragraphs.docx"));
        var s = new PowerNotes.Core.Edit.EditorSession(d);
        s.InsertFootnote("x");
        ValidSave(d);
    }

    [Fact]
    public void EmptyDocument_NavigationAndDeletes_DoNotThrow()
    {
        var (d, s) = NewDoc();
        s.Backspace(); s.DeleteForward(); Assert.False(s.Undo()); Assert.False(s.Redo());
        s.MoveHorizontal(-1, false); s.MoveHorizontal(1, true, true); s.MoveVertical(1, false); s.MoveVertical(-1, true, 500); s.MoveLineEdge(true, false, true);
        s.SelectAll(); s.DeleteSelection(); s.Type(""); s.SoftBreak(); s.Enter();
        Assert.Null(s.Find("", true)); Assert.Equal(0, s.ReplaceAll("", "x", false));
        ValidSave(d);
    }

    [Fact]
    public void Type_ControlCharacters_StaysSavable()
    {
        var (d, s) = NewDoc();
        s.Type("a\u0001b\u0000c￾d​e");
        ValidSave(d);
    }

    [Fact]
    public void ReplaceAll_WithTextContainingPattern_Terminates()
    {
        var (d, s) = NewDoc();
        s.Type("aaa");
        Assert.Equal(3, s.ReplaceAll("a", "aa", true));
        Assert.Equal("aaaaaa", ParaText.Text(d.Paragraphs()[0]));
    }

    [Fact]
    public void GarbageInputs_AreRejectedNotCrashing()
    {
        var (d, s) = NewDoc();
        s.Type("x");
        try { s.PasteFragment("<not-xml", true); } catch (Exception e) { Assert.IsNotType<NullReferenceException>(e); }
        s.SetSize(0); s.SetSize(-5); s.SetSize(double.NaN); Assert.Throws<ArgumentException>(() => s.SetColor("zzzz"));
        Assert.Throws<ArgumentOutOfRangeException>(() => s.SetColumns(99));
        s.SetIndent(double.NaN, 1e12, null);
        s.SetSpacing(-10, 1e9);
        ValidSave(d); d.Layout();
    }

    [Fact]
    public void DeleteAcrossFootnoteReference_KeepsDocumentConsistent()
    {
        var (d, s) = NewDoc();
        s.Type("ab"); s.InsertFootnote("n"); s.Type("cd");
        s.SelectAll(); s.DeleteSelection();
        d.Layout(); ValidSave(d);
        Assert.True(s.Undo());
        Assert.Equal("abcd", ParaText.Text(d.Paragraphs()[0]));
        Assert.Single(d.Body.Descendants<FootnoteReference>());
    }

    [Fact]
    public void ClearFormatting_And_BlankPage_AreUndoable()
    {
        var (d, s) = NewDoc();
        s.Type("grassetto"); s.SelectAll(); s.ToggleBold();
        Assert.True(d.Body.Descendants<Bold>().Any());
        s.ClearFormatting();
        Assert.False(d.Body.Descendants<Bold>().Any());
        Assert.True(s.Undo()); Assert.True(d.Body.Descendants<Bold>().Any());
        s.ClearSel(); s.InsertBlankPage();
        Assert.Equal(2, d.Body.Descendants<Break>().Count(b => b.Type?.Value == BreakValues.Page));
        Assert.True(d.Layout().Pages.Count >= 3);
        Assert.True(s.Undo());
        Assert.Empty(d.Body.Descendants<Break>());
        ValidSave(d);
    }
}
