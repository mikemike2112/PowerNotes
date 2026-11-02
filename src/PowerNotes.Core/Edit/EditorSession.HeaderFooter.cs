using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Model;

namespace PowerNotes.Core.Edit;

/// <summary>Contenuto di un'intestazione o piè di pagina nei termini della finestra di dialogo. Complex = contiene altro (più paragrafi, tabulazioni, altri campi, testo dopo il numero…): la finestra non può rappresentarlo senza perdere dati.</summary>
public sealed record HeaderFooterInfo(string Text, Align Align, PageNumFormat Pn, bool Complex);

public sealed partial class EditorSession
{
    /// <summary>Legge intestazione/piè di pagina della sezione del cursore: testo fisso (senza i campi), allineamento e tipo di numero di pagina (campi PAGE / NUMPAGES conservati come tali).</summary>
    public HeaderFooterInfo GetHeaderFooter(bool header)
    {
        var part = FindHF(SectionFor(Caret.P), header);
        OpenXmlElement? root = header ? (part as HeaderPart)?.Header : (OpenXmlElement?)(part as FooterPart)?.Footer;
        var paras = root?.Elements<Paragraph>().ToList() ?? [];
        if (root == null || paras.Count == 0) return new HeaderFooterInfo("", Align.Center, PageNumFormat.None, false);
        var al = paras[0].ParagraphProperties?.Justification?.Val?.Value;
        var align = al == JustificationValues.Center ? Align.Center : al == JustificationValues.Right ? Align.Right : Align.Left;
        string all = string.Join("\n", paras.Select(ParaText.Text));
        bool complex = paras.Count > 1 || root.ChildElements.Any(e => e is not (Paragraph or SectionProperties)) ||
                       paras[0].Descendants().Any(e => e is Drawing or Hyperlink or SimpleField or SdtRun or BookmarkStart or Picture or EmbeddedObject);
        var leaves = ParaText.Leaves(paras[0]);
        if (leaves.Any(l => l.Kind != LeafKind.Text || (l.Field != null && !IsPageField(l.Field)))) complex = true;
        // un campo fuori dai nostri due (es. DATE, TIME) senza risultato visibile non genera foglie: controlla anche le istruzioni
        if (paras[0].Descendants<FieldCode>().Any(f => !IsPageField(f.Text.Trim()))) complex = true;
        string t1 = "", t2 = "", t3 = ""; bool hasPage = false, hasTotal = false;
        foreach (var l in leaves)
        {
            if (l.Field != null) { if (l.Field.StartsWith("NUMPAGES", StringComparison.OrdinalIgnoreCase)) hasTotal = true; else hasPage = true; continue; }
            if (!hasPage && !hasTotal) t1 += l.Text; else if (hasPage && !hasTotal) t2 += l.Text; else t3 += l.Text;
        }
        var pn = hasPage && hasTotal ? PageNumFormat.PageOfTotal : hasPage ? PageNumFormat.Number : hasTotal ? PageNumFormat.None : PageNumFormat.None;
        if (hasTotal && !hasPage) complex = true;
        string text = t1;
        if (pn == PageNumFormat.PageOfTotal)
        {
            if (t2 != " di " || t3.Length > 0) complex = true;
            if (t1 == "Pagina ") text = ""; else if (t1.EndsWith(" Pagina ", StringComparison.Ordinal)) text = t1[..^8]; else complex = true;
        }
        else if (pn == PageNumFormat.Number && (t2.Length > 0 || t3.Length > 0)) complex = true;
        return new HeaderFooterInfo(complex ? all : text, align, pn, complex);
    }

    static bool IsPageField(string instr) { var f = instr.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.ToUpperInvariant(); return f is "PAGE" or "NUMPAGES"; }

    /// <summary>Numero di pagina nel piè di pagina («Pagina X di Y»), senza toccare testo e allineamento già presenti; se c'è già non fa nulla (nessun duplicato). Piè con contenuto complesso: rifiutato, va modificato sul foglio.</summary>
    public void InsertPageNumber()
    {
        var m = GetHeaderFooter(false);
        if (m.Complex) throw new InvalidOperationException("Il piè di pagina ha un contenuto complesso: modificalo direttamente sul foglio (doppio clic) per aggiungere il numero di pagina.");
        if (m.Pn != PageNumFormat.None) return;
        SetHeaderFooter(false, m.Text, m.Align, PageNumFormat.PageOfTotal);
    }

    /// <summary>Nome del formato carta (A4, A5, Letter, Legal) dalle due dimensioni in punti, indipendentemente dall'orientamento; «Personalizzato» altrimenti.</summary>
    public static string PaperName(double widthPt, double heightPt)
    {
        double a = Math.Min(widthPt, heightPt), b = Math.Max(widthPt, heightPt);
        bool Is(double w, double h) => Math.Abs(a - w) < 3 && Math.Abs(b - h) < 3;
        return Is(595.3, 841.9) ? "A4" : Is(419.5, 595.3) ? "A5" : Is(612, 792) ? "Letter" : Is(612, 1008) ? "Legal" : "Personalizzato";
    }
}

/// <summary>Pagine da stampare (indici da 0) per intervallo e copie scelti nella finestra di stampa.</summary>
public static class PrintPlan
{
    /// <param name="from">prima pagina (da 1)</param><param name="to">ultima pagina (da 1)</param><param name="collated">true: 1,2,3,1,2,3; false: 1,1,2,2,3,3</param>
    public static List<int> Pages(int total, int from, int to, int copies, bool collated)
    {
        var res = new List<int>();
        from = Math.Max(1, from); to = Math.Min(total, to); copies = Math.Clamp(copies, 1, 999);
        if (total < 1 || from > to) return res;
        if (collated) for (int c = 0; c < copies; c++) for (int p = from; p <= to; p++) res.Add(p - 1);
        else for (int p = from; p <= to; p++) for (int c = 0; c < copies; c++) res.Add(p - 1);
        return res;
    }
}
