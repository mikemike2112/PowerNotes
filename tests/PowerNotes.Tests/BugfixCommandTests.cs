using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Model;

namespace PowerNotes.Tests;

/// <summary>Correzioni dell'audit: numero di pagina duplicato, campi persi nel dialogo intestazione/piè, Legal/Letter, intervallo e copie di stampa.</summary>
public class BugfixCommandTests
{
    static (WordDoc d, EditorSession s) NewDoc() => TableCommandTests.NewDoc();
    static int Snap(WordDoc d) => d.Pkg.MainDocumentPart!.FooterParts.Select(f => f.Footer!.OuterXml).DefaultIfEmpty("").Aggregate((a, b) => a + b).GetHashCode();
    static List<string> Fields(WordDoc d) => d.Pkg.MainDocumentPart!.FooterParts.SelectMany(f => f.Footer!.Descendants<FieldCode>()).Select(f => f.Text.Trim()).ToList();

    [Fact]
    public void InsertPageNumber_Repeated_DoesNotDuplicate()
    {
        var (d, s) = NewDoc();
        s.Type("corpo");
        s.SetHeaderFooter(false, "Riservato", Align.Right, PageNumFormat.None);
        s.InsertPageNumber(); s.InsertPageNumber(); s.InsertPageNumber();
        Assert.Equal(["PAGE", "NUMPAGES"], Fields(d));
        var m = s.GetHeaderFooter(false);
        Assert.Equal(new HeaderFooterInfo("Riservato", Align.Right, PageNumFormat.PageOfTotal, false), m);
        Assert.Equal("Riservato Pagina 1 di 1", string.Join(" ", d.Layout().Pages[0].Prims.OfType<PowerNotes.Core.Layout.TextPrim>().Select(t => t.Text).Skip(1)));
        TableCommandTests.AssertValid(d);
    }

    [Fact]
    public void InsertPageNumber_EmptyFooter_DefaultsToCentered()
    {
        var (d, s) = NewDoc();
        s.Type("x");
        s.InsertPageNumber();
        Assert.Equal(new HeaderFooterInfo("", Align.Center, PageNumFormat.PageOfTotal, false), s.GetHeaderFooter(false));
        Assert.True(s.Undo());
        Assert.Empty(d.Body.Descendants<FooterReference>());
    }

    [Theory]
    [InlineData("", PageNumFormat.PageOfTotal, Align.Center)]
    [InlineData("Doc X", PageNumFormat.PageOfTotal, Align.Right)]
    [InlineData("Pag.", PageNumFormat.Number, Align.Left)]
    [InlineData("Solo testo", PageNumFormat.None, Align.Left)]
    public void HeaderFooterDialog_PrefillThenSave_KeepsFields(string text, PageNumFormat pn, Align al)
    {
        var (d, s) = NewDoc();
        s.Type("x");
        s.SetHeaderFooter(false, text, al, pn);
        var before = Fields(d);
        var m = s.GetHeaderFooter(false);
        Assert.Equal(new HeaderFooterInfo(text, al, pn, false), m);
        s.SetHeaderFooter(false, m.Text, m.Align, m.Pn);          // OK senza modifiche nel dialogo
        Assert.Equal(before, Fields(d));
        s.SetHeaderFooter(false, m.Text, m.Align, m.Pn);
        Assert.Equal(before, Fields(d));
        TableCommandTests.AssertValid(d);
    }

    [Fact]
    public void ComplexFooter_IsDetected_AndNeverRewritten()
    {
        var (d, s) = NewDoc();
        s.Type("x");
        s.SetHeaderFooter(false, "uno", Align.Left, PageNumFormat.None);
        var ft = d.Pkg.MainDocumentPart!.FooterParts.Single().Footer!;
        ft.Append(new Paragraph(new Run(new Text("seconda riga"))));
        Assert.True(s.GetHeaderFooter(false).Complex);
        int before = Snap(d);
        Assert.Throws<InvalidOperationException>(() => s.InsertPageNumber());
        Assert.Equal(before, Snap(d));
        // campo DATE e testo dopo il numero: complessi
        ft.RemoveAllChildren();
        ft.Append(new Paragraph(new SimpleField { Instruction = "DATE", }, new Run(new Text("fine"))));
        Assert.True(s.GetHeaderFooter(false).Complex);
        ft.RemoveAllChildren();
        ft.Append(new Paragraph(new Run(new Text("a")), new Run(new TabChar()), new Run(new Text("b"))));
        Assert.True(s.GetHeaderFooter(false).Complex);
    }

    [Theory]
    [InlineData(612, 792, "Letter")]
    [InlineData(612, 1008, "Legal")]
    [InlineData(1008, 612, "Legal")]
    [InlineData(792, 612, "Letter")]
    [InlineData(595.3, 841.9, "A4")]
    [InlineData(841.9, 595.3, "A4")]
    [InlineData(419.5, 595.3, "A5")]
    [InlineData(500, 700, "Personalizzato")]
    public void PaperName_UsesBothDimensions(double w, double h, string expected) => Assert.Equal(expected, EditorSession.PaperName(w, h));

    [Fact]
    public void PrintPlan_RangeAndCopies()
    {
        Assert.Equal([0, 1, 2, 3], PrintPlan.Pages(4, 1, 4, 1, true));
        Assert.Equal([1, 2], PrintPlan.Pages(4, 2, 3, 1, true));
        Assert.Equal([1, 2, 1, 2], PrintPlan.Pages(4, 2, 3, 2, true));
        Assert.Equal([1, 1, 2, 2], PrintPlan.Pages(4, 2, 3, 2, false));
        Assert.Equal([2, 3], PrintPlan.Pages(4, 3, 99, 1, true));        // oltre la fine: ridotto
        Assert.Equal([0], PrintPlan.Pages(4, -5, 1, 1, true));
        Assert.Empty(PrintPlan.Pages(4, 3, 2, 1, true));
        Assert.Empty(PrintPlan.Pages(0, 1, 1, 1, true));
        Assert.Equal(999, PrintPlan.Pages(1, 1, 1, 100000, true).Count);
    }
}
