using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Model;

namespace PowerNotes.Core.Edit;

/// <summary>Operazioni elementari sull'XML di un paragrafo, indirizzate per offset di carattere. Toccano solo gli elementi coinvolti.</summary>
public static class XmlOps
{
    static IEnumerable<OpenXmlElement> Content(Run r) => r.ChildElements.Where(c => c is not RunProperties);

    static Run NewRunLike(Run r)
    {
        var nr = (Run)r.CloneNode(false);
        if (r.RunProperties != null) nr.Append(r.RunProperties.CloneNode(true));
        return nr;
    }

    static void SplitText(Leaf l, int k)
    {
        var t = (Text)l.El; var r = l.Run;
        var nr = NewRunLike(r);
        var rest = new Text(t.Text[k..]) { Space = SpaceProcessingModeValues.Preserve };
        t.Text = t.Text[..k]; t.Space = SpaceProcessingModeValues.Preserve;
        nr.Append(rest);
        foreach (var sib in t.ElementsAfter().ToList()) { sib.Remove(); nr.Append(sib); }
        r.InsertAfterSelf(nr);
    }

    static void SplitRunBefore(Leaf l)
    {
        var r = l.Run;
        if (!l.El.ElementsBefore().Any(c => c is not RunProperties)) return;
        var nr = NewRunLike(r);
        var move = new List<OpenXmlElement> { l.El };
        move.AddRange(l.El.ElementsAfter());
        foreach (var m in move) { m.Remove(); nr.Append(m); }
        r.InsertAfterSelf(nr);
    }

    /// <summary>Garantisce un confine di run esattamente all'offset.</summary>
    public static void EnsureBoundary(Paragraph p, int off)
    {
        var leaves = ParaText.Leaves(p);
        var inside = leaves.FirstOrDefault(l => l.Kind == LeafKind.Text && l.Start < off && off < l.End);
        if (inside != null) { SplitText(inside, off - inside.Start); leaves = ParaText.Leaves(p); }
        var at = leaves.FirstOrDefault(l => l.Start == off);
        if (at != null) SplitRunBefore(at);
    }

    static RunProperties? MarkAsRun(Paragraph p)
    {
        var m = p.ParagraphProperties?.ParagraphMarkRunProperties;
        if (m == null) return null;
        var rp = new RunProperties();
        foreach (var c in m.ChildElements) if (c is not (Inserted or Deleted or MoveFrom or MoveTo)) rp.Append(c.CloneNode(true));
        return rp;
    }

    /// <summary>Toglie i caratteri non ammessi in XML 1.0 (controlli, U+FFFE/FFFF): salvandoli il file non si scriverebbe più.</summary>
    public static string Clean(string t) => System.Text.RegularExpressions.Regex.Replace(t, @"[^\u0009\u000A\u000D\u0020-\uFFFD]", "");

    public static Run MakeRun(string text, RunProperties? rpr)
    {
        text = Clean(text);
        var r = new Run();
        if (rpr != null) r.Append(rpr.CloneNode(true));
        r.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return r;
    }

    /// <summary>Inserisce testo (senza a capo né tab) all'offset. Eredita la formattazione del carattere precedente, o quella indicata.</summary>
    public static void InsertText(Paragraph p, int off, string text, RunProperties? forced = null)
    {
        if (text.Length == 0) return;
        var leaves = ParaText.Leaves(p);
        if (forced == null)
        {
            var tl = leaves.LastOrDefault(l => l.Kind == LeafKind.Text && l.Start < off && off <= l.End)
                  ?? leaves.FirstOrDefault(l => l.Kind == LeafKind.Text && l.Start == off && off == 0);
            if (tl != null && tl.Field == null)
            {
                var t = (Text)tl.El;
                int k = off - tl.Start;
                t.Text = t.Text.Insert(k, text); t.Space = SpaceProcessingModeValues.Preserve;
                return;
            }
        }
        var rpr = forced ?? (leaves.LastOrDefault(l => l.Start < off && l.Kind != LeafKind.Image)?.Run.RunProperties
                          ?? leaves.FirstOrDefault(l => l.Start >= off && l.Kind != LeafKind.Image)?.Run.RunProperties ?? MarkAsRun(p));
        PlaceRun(p, off, MakeRun(text, rpr));
    }

    /// <summary>Inserisce un run già pronto all'offset (spezzando i run se serve).</summary>
    public static void PlaceRun(Paragraph p, int off, Run run)
    {
        EnsureBoundary(p, off);
        var at = ParaText.Leaves(p).FirstOrDefault(l => l.Start == off);
        if (at != null) at.Run.InsertBeforeSelf(run);
        else
        {
            var lastRun = ParaText.Leaves(p).LastOrDefault()?.Run;
            if (lastRun != null) lastRun.InsertAfterSelf(run); else p.Append(run);
        }
    }

    public static void InsertTab(Paragraph p, int off)
    {
        var leaves = ParaText.Leaves(p);
        var rpr = leaves.LastOrDefault(l => l.Start < off && l.Kind != LeafKind.Image)?.Run.RunProperties ?? MarkAsRun(p);
        var r = new Run(); if (rpr != null) r.Append(rpr.CloneNode(true)); r.Append(new TabChar());
        PlaceRun(p, off, r);
    }

    public static void InsertBreak(Paragraph p, int off, BreakValues? type)
    {
        var leaves = ParaText.Leaves(p);
        var rpr = leaves.LastOrDefault(l => l.Start < off && l.Kind != LeafKind.Image)?.Run.RunProperties ?? MarkAsRun(p);
        var r = new Run(); if (rpr != null) r.Append(rpr.CloneNode(true));
        r.Append(type == null ? new Break() : new Break { Type = type });
        PlaceRun(p, off, r);
    }

    public static void DeleteRange(Paragraph p, int a, int b)
    {
        if (b <= a) return;
        EnsureBoundary(p, b); EnsureBoundary(p, a);
        var gone = ParaText.Leaves(p).Where(l => l.Start >= a && l.End <= b).ToList();
        var runs = new HashSet<Run>();
        foreach (var l in gone) { runs.Add(l.Run); l.El.Remove(); }
        foreach (var r in runs) if (!Content(r).Any()) r.Remove();
    }

    /// <summary>Divide il paragrafo all'offset; restituisce il secondo (che eredita pPr, compresa sectPr).</summary>
    public static Paragraph Split(Paragraph p, int off)
    {
        EnsureBoundary(p, off);
        var np = new Paragraph();
        if (p.ParagraphProperties != null)
        {
            np.Append(p.ParagraphProperties.CloneNode(true));
            p.ParagraphProperties.SectionProperties?.Remove();
            np.ParagraphProperties!.PageBreakBefore = null;
        }
        var at = ParaText.Leaves(p).FirstOrDefault(l => l.Start == off);
        if (at != null)
        {
            OpenXmlElement node = at.Run;
            while (node.Parent != p)
            {
                var parent = node.Parent!;
                var clone = parent.CloneNode(false);
                var move = new List<OpenXmlElement> { node };
                move.AddRange(node.ElementsAfter());
                foreach (var m in move) { m.Remove(); clone.Append(m); }
                parent.InsertAfterSelf(clone);
                node = clone;
            }
            var tail = new List<OpenXmlElement> { node };
            tail.AddRange(node.ElementsAfter().Where(c => c is not ParagraphProperties));
            foreach (var m in tail) { m.Remove(); np.Append(m); }
        }
        p.InsertAfterSelf(np);
        return np;
    }

    /// <summary>Unisce next in p (p mantiene le proprietà).</summary>
    public static void Merge(Paragraph p, Paragraph next)
    {
        foreach (var c in next.ChildElements.Where(c => c is not ParagraphProperties).ToList()) { c.Remove(); p.Append(c); }
        // se p era vuoto e next ha sectPr, trasferiscila
        if (next.ParagraphProperties?.SectionProperties is { } sp && p.ParagraphProperties?.SectionProperties == null)
        {
            p.ParagraphProperties ??= new ParagraphProperties();
            p.ParagraphProperties.SectionProperties = (SectionProperties)sp.CloneNode(true);
        }
        next.Remove();
    }

    public static RunProperties RPr(Run r) => r.RunProperties ?? r.PrependChild(new RunProperties());

    /// <summary>Applica una modifica alle proprietà dei run compresi in [a,b).</summary>
    public static void Format(Paragraph p, int a, int b, Action<RunProperties> f)
    {
        if (b <= a) return;
        EnsureBoundary(p, b); EnsureBoundary(p, a);
        foreach (var r in ParaText.Leaves(p).Where(l => l.Start >= a && l.End <= b).Select(l => l.Run).Distinct().ToList()) f(RPr(r));
    }

    public static ParagraphProperties PPr(Paragraph p) => p.ParagraphProperties ?? p.PrependChild(new ParagraphProperties());
}
