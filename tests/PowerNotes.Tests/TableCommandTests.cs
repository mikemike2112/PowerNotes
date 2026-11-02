using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Model;

namespace PowerNotes.Tests;

/// <summary>Comandi della scheda «Layout di tabella»: effetto, un solo annulla, file valido per il validatore Open XML.</summary>
public class TableCommandTests
{
    internal static (WordDoc d, EditorSession s) NewDoc() { var d = WordDoc.CreateNew(); return (d, new EditorSession(d)); }

    /// <summary>Salva, riapre e valida con il validatore Open XML; restituisce il documento riletto.</summary>
    internal static void AssertValid(WordDoc d)
    {
        var path = Path.Combine(Path.GetTempPath(), "pn-cmd-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            d.Save(path);
            using var pk = WordprocessingDocument.Open(path, false);
            var errs = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(pk).Select(e => e.Description + " @ " + e.Path?.XPath).ToList();
            Assert.True(errs.Count == 0, string.Join(" | ", errs));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    static (WordDoc d, EditorSession s, Table t) Tbl(int rows, int cols)
    {
        var (d, s) = NewDoc();
        s.InsertTable(rows, cols);
        var t = d.Body.Descendants<Table>().First();
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                var p = t.Elements<TableRow>().ElementAt(r).Elements<TableCell>().ElementAt(c).Elements<Paragraph>().First();
                s.SetCaret(p, 0); s.Type($"r{r}c{c}");
            }
        s.SetCaret(t.Descendants<Paragraph>().First(), 0);
        return (d, s, t);
    }

    static int Snap(WordDoc d) => d.Body.OuterXml.GetHashCode();

    static void CaretIn(EditorSession s, Table t, int r, int c) => s.SetCaret(t.Elements<TableRow>().ElementAt(r).Elements<TableCell>().ElementAt(c).Elements<Paragraph>().First(), 0);

    [Fact]
    public void SplitCell_AddsColumnsAndUndoesInOneStep()
    {
        var (d, s, t) = Tbl(2, 2);
        int before = Snap(d);
        s.SplitCell(3);
        Assert.Equal(4, t.Elements<TableRow>().First().Elements<TableCell>().Count());
        Assert.Equal(2, t.Elements<TableRow>().Last().Elements<TableCell>().Count());
        Assert.Equal(4, t.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
        Assert.Equal(3, t.Elements<TableRow>().Last().Elements<TableCell>().First().TableCellProperties!.GridSpan!.Val!.Value);
        AssertValid(d);
        Assert.True(s.Undo());
        Assert.Equal(before, Snap(d));
    }

    [Fact]
    public void SplitCell_OnMergedCellDividesSpan()
    {
        var (d, s, t) = Tbl(2, 3);
        var ps = t.Elements<TableRow>().First().Descendants<Paragraph>().ToList();
        s.SetCaret(ps[0], 0); s.SetCaret(ps[1], 0, true);
        Assert.True(s.MergeCells());
        CaretIn(s, t, 0, 0);
        s.SplitCell(2);
        Assert.Equal(3, t.Elements<TableRow>().First().Elements<TableCell>().Count());
        Assert.Equal(3, t.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
        AssertValid(d);
    }

    [Fact]
    public void SplitTable_SeparatesRowsWithParagraph()
    {
        var (d, s, t) = Tbl(3, 2);
        CaretIn(s, t, 1, 0);
        s.SplitTable();
        var tables = d.Body.Descendants<Table>().ToList();
        Assert.Equal(2, tables.Count);
        Assert.Single(tables[0].Elements<TableRow>()); Assert.Equal(2, tables[1].Elements<TableRow>().Count());
        Assert.IsType<Paragraph>(tables[0].NextSibling());
        AssertValid(d);
        CaretIn(s, tables[0], 0, 0);
        Assert.Throws<InvalidOperationException>(() => s.SplitTable());
        Assert.True(s.Undo());
        Assert.Single(d.Body.Descendants<Table>());
    }

    [Fact]
    public void DistributeRowsAndColumns_MakeThemEqual()
    {
        var (d, s, t) = Tbl(3, 3);
        var g = t.GetFirstChild<TableGrid>()!.Elements<GridColumn>().ToList();
        g[0].Width = "1000"; g[1].Width = "2000"; g[2].Width = "3000";
        s.DistributeColumns();
        var w = t.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Select(x => int.Parse(x.Width!.Value!)).ToList();
        Assert.Equal(6000, w.Sum()); Assert.Equal(w[0], w[1]);
        Assert.Equal(w[0].ToString(), t.Descendants<TableCell>().First().TableCellProperties!.TableCellWidth!.Width!.Value);
        s.DistributeRows();
        var hs = t.Elements<TableRow>().Select(r => r.TableRowProperties!.GetFirstChild<TableRowHeight>()!.Val!.Value).Distinct().ToList();
        Assert.Single(hs);
        AssertValid(d);
        Assert.True(s.Undo()); Assert.Null(d.Body.Descendants<Table>().First().Elements<TableRow>().First().TableRowProperties?.GetFirstChild<TableRowHeight>());
    }

    [Fact]
    public void AutoFit_ModesSetWidthsAndValidate()
    {
        var (d, s, t) = Tbl(2, 3);
        s.AutoFit(TableFit.Window);
        var tp = t.GetFirstChild<TableProperties>()!;
        Assert.Equal("5000", tp.TableWidth!.Width!.Value);
        Assert.InRange(t.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Sum(x => int.Parse(x.Width!.Value!)), 9000, 11000);
        s.AutoFit(TableFit.Fixed);
        Assert.Equal(TableLayoutValues.Fixed, tp.TableLayout!.Type!.Value);
        s.AutoFit(TableFit.Content);
        Assert.Null(tp.TableLayout);
        Assert.True(t.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Sum(x => int.Parse(x.Width!.Value!)) < 4000);
        AssertValid(d);
    }

    [Fact]
    public void SortTable_NumericAndText_KeepsHeader()
    {
        var (d, s, t) = Tbl(4, 2);
        string[] vals = ["Valore", "10", "9", "100"];
        for (int r = 0; r < 4; r++) { var p = t.Elements<TableRow>().ElementAt(r).Elements<TableCell>().Last().Elements<Paragraph>().First(); s.SetCaret(p, 0); s.SetCaret(p, ParaText.Length(p), true); s.Type(vals[r]); }
        s.ToggleRepeatHeader();   // la prima riga è l'intestazione
        CaretIn(s, t, 1, 1);
        s.SortTable(false);
        string Col(int c) => string.Join(",", d.Body.Descendants<Table>().First().Elements<TableRow>().Select(r => ParaText.Text(r.Elements<TableCell>().ElementAt(c).Elements<Paragraph>().First())));
        Assert.Equal("Valore,9,10,100", Col(1));
        s.SortTable(true);
        Assert.Equal("Valore,100,10,9", Col(1));
        Assert.StartsWith("r0c0,", Col(0));
        AssertValid(d);
        Assert.True(s.Undo()); Assert.Equal("Valore,9,10,100", Col(1));
    }

    [Fact]
    public void SortTable_MergedCells_Refused()
    {
        var (d, s, t) = Tbl(2, 2);
        var ps = t.Elements<TableRow>().First().Descendants<Paragraph>().ToList();
        s.SetCaret(ps[0], 0); s.SetCaret(ps[1], 0, true); s.MergeCells();
        CaretIn(s, t, 1, 0);
        Assert.Throws<InvalidOperationException>(() => s.SortTable(false));
    }

    [Fact]
    public void RepeatHeader_TogglesAndSurvivesRoundTrip()
    {
        var (d, s, t) = Tbl(3, 2);
        Assert.True(s.ToggleRepeatHeader());
        Assert.NotNull(t.Elements<TableRow>().First().TableRowProperties!.GetFirstChild<TableHeader>());
        Assert.Null(t.Elements<TableRow>().Last().TableRowProperties?.GetFirstChild<TableHeader>());
        AssertValid(d);
        Assert.False(s.ToggleRepeatHeader());
        Assert.False(s.HasRepeatHeader);
    }

    [Fact]
    public void ConvertTableToText_KeepsTextWithTabs()
    {
        var (d, s, t) = Tbl(2, 2);
        int before = Snap(d);
        s.ConvertTableToText();
        Assert.Empty(d.Body.Descendants<Table>());
        var texts = d.Paragraphs().Select(ParaText.Text).Where(x => x.Length > 0).ToList();
        Assert.Equal(["r0c0\tr0c1", "r1c0\tr1c1"], texts);
        AssertValid(d);
        Assert.True(s.Undo()); Assert.Equal(before, Snap(d));
    }

    [Fact]
    public void SelectTablePart_SelectsRowColumnTable()
    {
        var (d, s, t) = Tbl(3, 3);
        CaretIn(s, t, 1, 1);
        s.SelectTablePart(TableSel.Row);
        Assert.Equal("r1c0\nr1c1\nr1c2", s.SelectedText().Replace("\n\n", "\n"));
        CaretIn(s, t, 1, 1);
        s.SelectTablePart(TableSel.Column);
        Assert.Contains("r0c1", s.SelectedText()); Assert.Contains("r2c1", s.SelectedText());
        CaretIn(s, t, 1, 1);
        s.SelectTablePart(TableSel.Table);
        Assert.Contains("r0c0", s.SelectedText()); Assert.Contains("r2c2", s.SelectedText());
        CaretIn(s, t, 1, 1);
        s.SelectTablePart(TableSel.Cell);
        Assert.Equal("r1c1", s.SelectedText());
    }

    [Fact]
    public void TableProps_AlignWidthRowHeight_AndLayoutCenters()
    {
        var (d, s, t) = Tbl(2, 2);
        s.AutoFit(TableFit.Content);
        double x0 = s.Layout.LinesOf(t.Descendants<Paragraph>().First())[0].Segs[0].X;
        s.SetTableProps("Centro", 50, 30);
        var pr = s.GetTableProps();
        Assert.Equal("Centro", pr.Align); Assert.Equal(50, pr.WidthPct); Assert.Equal(30, pr.RowHeightPt);
        Assert.NotNull(t.GetFirstChild<TableProperties>()!.TableJustification);
        double x1 = s.Layout.LinesOf(t.Descendants<Paragraph>().First())[0].Segs[0].X;
        Assert.True(x1 > x0 + 50, $"{x0} -> {x1}");
        AssertValid(d);
        Assert.Throws<ArgumentOutOfRangeException>(() => s.SetTableProps("Centro", 101, 0));
        s.SetTableProps("Sinistra", 0, 0);
        Assert.Null(t.GetFirstChild<TableProperties>()!.TableJustification);
        Assert.Null(t.Elements<TableRow>().First().TableRowProperties?.GetFirstChild<TableRowHeight>());
    }

    [Fact]
    public void Layout_RecordsCellRectsForGridView()
    {
        var (d, s, t) = Tbl(2, 2);
        Assert.Equal(4, s.Layout.Pages[0].Cells.Count);
    }

    [Fact]
    public void NotInTable_Throws()
    {
        var (d, s) = NewDoc();
        Assert.Throws<InvalidOperationException>(() => s.SplitCell(2));
        Assert.Throws<InvalidOperationException>(() => s.DistributeRows());
        Assert.Throws<InvalidOperationException>(() => s.SortTable(false));
    }
}
