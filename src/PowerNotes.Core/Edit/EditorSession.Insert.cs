using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Model;

namespace PowerNotes.Core.Edit;

/// <summary>Scheda Inserisci / Riferimenti: collegamenti, segnalibri, note di chiusura, navigazione fra le note.</summary>
public sealed partial class EditorSession
{
    // ---------- collegamenti ----------
    static readonly string[] LinkSchemes = ["http", "https", "mailto", "ftp"];

    /// <summary>Normalizza e valida l'indirizzo di un collegamento: solo http/https/mailto/ftp (mai file:, javascript: ecc.) oppure «#segnalibro».</summary>
    public static string NormalizeLink(string target)
    {
        target = (target ?? "").Trim();
        if (target.Length == 0 || target.Any(char.IsControl) || target.Length > 2000) throw new ArgumentException("Indirizzo del collegamento non valido.");
        if (target[0] == '#') return target.Length > 1 ? target : throw new ArgumentException("Indirizzo del collegamento non valido.");
        if (!target.Contains(':') && target.Contains('@') && !target.Contains('/')) target = "mailto:" + target;
        else if (!Regex.IsMatch(target, @"^[A-Za-z][A-Za-z0-9+.\-]*:")) target = "https://" + target;
        if (!Uri.TryCreate(target, UriKind.Absolute, out var u) || !LinkSchemes.Contains(u.Scheme.ToLowerInvariant()) || (u.Scheme != "mailto" && u.Host.Length == 0))
            throw new ArgumentException("Indirizzo del collegamento non valido: usa http, https, mailto o ftp.");
        return target;
    }

    Hyperlink? LinkAtCaret()
    {
        var leaves = ParaText.Leaves(Caret.P);
        var leaf = leaves.LastOrDefault(l => l.Start < Caret.Off && Caret.Off <= l.End) ?? leaves.FirstOrDefault(l => l.Start == Caret.Off);
        return leaf?.Run.Parent as Hyperlink;
    }

    string? TargetOf(Hyperlink h)
    {
        if (h.Anchor?.Value is { Length: > 0 } a) return "#" + a;
        var id = h.Id?.Value;
        return id == null ? null : Doc.Pkg.MainDocumentPart!.HyperlinkRelationships.FirstOrDefault(r => r.Id == id)?.Uri.OriginalString;
    }

    /// <summary>Indirizzo del collegamento sotto il cursore («#nome» per i segnalibri), altrimenti null.</summary>
    public string? HyperlinkAt() => LinkAtCaret() is { } h ? TargetOf(h) : null;

    void SetTarget(Hyperlink h, string target)
    {
        var main = Doc.Pkg.MainDocumentPart!;
        string? old = h.Id?.Value;
        if (target[0] == '#') { h.Id = null; h.Anchor = target[1..]; }
        else { h.Anchor = null; h.Id = main.AddHyperlinkRelationship(new Uri(target, UriKind.Absolute), true).Id; }
        h.History = true;
        DropUnusedLinkRel(old);
    }

    void DropUnusedLinkRel(string? id)
    {
        if (id == null) return;
        var main = Doc.Pkg.MainDocumentPart!;
        if (!Doc.Body.Descendants<Hyperlink>().Any(x => x.Id?.Value == id) && main.HyperlinkRelationships.Any(r => r.Id == id)) main.DeleteExternalRelationship(id);
    }

    /// <summary>Collega la selezione (o, senza selezione, inserisce <paramref name="text"/> come collegamento). Se il cursore è già in un collegamento ne cambia l'indirizzo.</summary>
    public void SetHyperlink(string target, string? text = null)
    {
        target = NormalizeLink(target);
        if (Caret.P.Ancestors<Body>().Count() == 0) throw new InvalidOperationException("I collegamenti si inseriscono solo nel corpo del documento.");
        if (target[0] == '#' && !Doc.Body.Descendants<BookmarkStart>().Any(b => b.Name?.Value == target[1..])) throw new ArgumentException("Il segnalibro «" + target[1..] + "» non esiste.");
        var styleId = BuiltIns.EnsureStyle(Doc.Pkg, "Hyperlink");
        Begin();
        if (!HasSelection)
        {
            if (LinkAtCaret() is { } cur) { SetTarget(cur, target); Changed(); return; }
            text = XmlOps.Clean(string.IsNullOrWhiteSpace(text) ? (target[0] == '#' ? target[1..] : target) : text);
            var h = new Hyperlink();
            h.Append(new Run(new RunProperties(new RunStyle { Val = styleId }), new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            SetTarget(h, target);
            XmlOps.EnsureBoundary(Caret.P, Caret.Off);
            var at = ParaText.Leaves(Caret.P).FirstOrDefault(l => l.Start == Caret.Off);
            if (at != null) { OpenXmlElement top = at.Run; while (top.Parent != Caret.P) top = top.Parent!; top.InsertBeforeSelf(h); }
            else Caret.P.Append(h);
            Set(new Pos(Caret.P, Caret.Off + text.Length));
            Changed(); return;
        }
        var (s, e) = Range;
        // già collegamento: cambia solo l'indirizzo
        var existing = new List<Hyperlink>();
        var wrap = new List<(Paragraph P, List<Run> Runs)>();
        foreach (var p in SelectedParagraphs())
        {
            int a = p == s.P ? s.Off : 0, b = p == e.P ? e.Off : ParaText.Length(p);
            if (b <= a) continue;
            XmlOps.EnsureBoundary(p, b); XmlOps.EnsureBoundary(p, a);
            var runs = ParaText.Leaves(p).Where(l => l.Start >= a && l.End <= b).Select(l => l.Run).Distinct().ToList();
            if (runs.Any(r => r.Parent is Hyperlink)) { existing.AddRange(runs.Select(r => r.Parent).OfType<Hyperlink>()); continue; }
            if (runs.Any(r => r.Parent != p)) throw new InvalidOperationException("Il testo selezionato non può diventare un collegamento qui.");
            if (runs.Count > 0) wrap.Add((p, runs));
        }
        foreach (var h in existing.Distinct()) SetTarget(h, target);
        foreach (var (p, runs) in wrap)
        {
            var h = new Hyperlink();
            SetTarget(h, target);
            int i0 = p.ChildElements.ToList().IndexOf(runs[0]), i1 = p.ChildElements.ToList().IndexOf(runs[^1]);
            var span = p.ChildElements.Skip(i0).Take(i1 - i0 + 1).ToList();
            runs[0].InsertBeforeSelf(h);
            foreach (var x in span) { x.Remove(); h.Append(x); }
            foreach (var r in runs) XmlOps.RPr(r).RunStyle = new RunStyle { Val = styleId };
        }
        Changed();
    }

    /// <summary>Toglie il collegamento sotto il cursore (o nella selezione), lasciando il testo. True se c'era qualcosa da togliere.</summary>
    public bool RemoveHyperlink()
    {
        var links = new List<Hyperlink>();
        if (HasSelection)
        {
            var (s, e) = Range;
            foreach (var p in SelectedParagraphs())
            {
                int a = p == s.P ? s.Off : 0, b = p == e.P ? e.Off : ParaText.Length(p);
                links.AddRange(ParaText.Leaves(p).Where(l => l.End > a && l.Start < b).Select(l => l.Run.Parent).OfType<Hyperlink>());
            }
        }
        else if (LinkAtCaret() is { } h) links.Add(h);
        links = links.Distinct().ToList();
        if (links.Count == 0) return false;
        Begin();
        var styleId = BuiltIns.EnsureStyle(Doc.Pkg, "Hyperlink");
        foreach (var h in links)
        {
            string? id = h.Id?.Value;
            foreach (var ch in h.ChildElements.ToList())
            {
                ch.Remove(); h.InsertBeforeSelf(ch);
                if (ch is Run r && r.RunProperties?.RunStyle?.Val?.Value == styleId) { r.RunProperties.RunStyle = null; if (!r.RunProperties.HasChildren) r.RunProperties = null; }
            }
            h.Remove();
            DropUnusedLinkRel(id);
        }
        Changed();
        return true;
    }

    // ---------- segnalibri ----------
    static readonly Regex BookmarkName = new("^[A-Za-z][A-Za-z0-9_]{0,39}$");

    /// <summary>Segnalibri visibili (quelli che iniziano con «_» sono nascosti, come in Word), in ordine alfabetico.</summary>
    public List<string> Bookmarks() => Doc.Body.Descendants<BookmarkStart>().Select(b => b.Name?.Value ?? "").Where(n => n.Length > 0 && n[0] != '_').Distinct().OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();

    void PutMarker(Paragraph p, int off, OpenXmlElement marker)
    {
        XmlOps.EnsureBoundary(p, off);
        var at = ParaText.Leaves(p).FirstOrDefault(l => l.Start == off);
        if (at == null) { p.Append(marker); return; }
        OpenXmlElement top = at.Run; while (top.Parent != p) top = top.Parent!;
        top.InsertBeforeSelf(marker);
    }

    /// <summary>Aggiunge un segnalibro sulla selezione (o sul punto del cursore). Se il nome esiste già, il segnalibro si sposta qui.</summary>
    public void AddBookmark(string name)
    {
        name = (name ?? "").Trim();
        if (!BookmarkName.IsMatch(name)) throw new ArgumentException("Nome del segnalibro non valido: inizia con una lettera, solo lettere, cifre e _ (massimo 40 caratteri).");
        if (Caret.P.Ancestors<Body>().Count() == 0) throw new InvalidOperationException("I segnalibri si inseriscono solo nel corpo del documento.");
        Begin();
        RemoveBookmarkCore(name);
        long id = Doc.Body.Descendants<BookmarkStart>().Select(b => long.TryParse(b.Id?.Value, out var v) ? v : 0).DefaultIfEmpty(0).Max() + 1;
        var (s, e) = Range;
        PutMarker(s.P, s.Off, new BookmarkStart { Id = id.ToString(), Name = name });
        PutMarker(e.P, e.Off, new BookmarkEnd { Id = id.ToString() });
        Changed();
    }

    bool RemoveBookmarkCore(string name)
    {
        var st = Doc.Body.Descendants<BookmarkStart>().Where(b => b.Name?.Value == name).ToList();
        foreach (var b in st) { var id = b.Id?.Value; foreach (var en in Doc.Body.Descendants<BookmarkEnd>().Where(x => x.Id?.Value == id).ToList()) en.Remove(); b.Remove(); }
        return st.Count > 0;
    }

    public bool DeleteBookmark(string name)
    {
        if (!Doc.Body.Descendants<BookmarkStart>().Any(b => b.Name?.Value == name)) return false;
        Begin(); RemoveBookmarkCore(name); Changed(); return true;
    }

    /// <summary>Offset del paragrafo che corrisponde alla posizione di un elemento non testuale (segnalibro, richiamo di nota).</summary>
    static int OffsetBefore(Paragraph p, OpenXmlElement marker)
    {
        var all = p.Descendants().ToList();
        int idx = all.IndexOf(marker), off = 0;
        foreach (var l in ParaText.Leaves(p)) if (all.IndexOf(l.Run) < idx) off = l.End;
        return off;
    }

    public bool GoToBookmark(string name)
    {
        var b = Doc.Body.Descendants<BookmarkStart>().FirstOrDefault(x => x.Name?.Value == name);
        var p = b?.Ancestors<Paragraph>().FirstOrDefault();
        if (b == null || p == null || !Paragraphs().Contains(p)) return false;
        Set(new Pos(p, OffsetBefore(p, b)));
        return true;
    }

    // ---------- note di chiusura ----------
    /// <summary>Inserisce un richiamo di nota di chiusura al cursore e crea la nota (elencata in fondo al documento).</summary>
    public void InsertEndnote(string text)
    {
        if (Caret.P.Ancestors<Body>().Count() == 0 || Caret.P.Ancestors<Footnote>().Any() || Caret.P.Ancestors<Endnote>().Any()) throw new InvalidOperationException("Le note si inseriscono solo nel corpo del documento.");
        text = XmlOps.Clean(text ?? "");
        Begin();
        var main = Doc.Pkg.MainDocumentPart!;
        const string Wns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        var ep = main.EndnotesPart ?? main.AddNewPart<EndnotesPart>();
        ep.Endnotes ??= new Endnotes(
            new Endnote($"<w:endnote {Wns} w:type=\"separator\" w:id=\"-1\"><w:p><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr><w:r><w:separator/></w:r></w:p></w:endnote>"),
            new Endnote($"<w:endnote {Wns} w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>"));
        if (main.DocumentSettingsPart?.Settings is { } st && st.GetFirstChild<EndnoteDocumentWideProperties>() == null)
        {
            var after = st.ChildElements.FirstOrDefault(x => x.LocalName is "compat" or "docVars" or "rsids" or "mathPr" or "attachedSchema" or "themeFontLang" or "clrSchemeMapping" or "doNotIncludeSubdocsInStats" or "doNotAutoCompressPictures" or "forceUpgrade" or "captions" or "readModeInkLockDown" or "smartTagType" or "schemaLibrary" or "shapeDefaults" or "doNotEmbedSmartTags" or "decimalSymbol" or "listSeparator");
            var epr = new EndnoteDocumentWideProperties(new EndnoteSpecialReference { Id = -1 }, new EndnoteSpecialReference { Id = 0 });
            if (after != null) st.InsertBefore(epr, after); else st.Append(epr);
        }
        var sty = BuiltIns.EnsureStyle(Doc.Pkg, "EndnoteText"); var rsty = BuiltIns.EnsureStyle(Doc.Pkg, "EndnoteReference");
        long id = ep.Endnotes.Elements<Endnote>().Select(f => f.Id?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
        XmlOps.PlaceRun(Caret.P, Caret.Off, new Run(new RunProperties(new RunStyle { Val = rsty }), new EndnoteReference { Id = id }));
        var np = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = sty }),
            new Run(new RunProperties(new RunStyle { Val = rsty }), new EndnoteReferenceMark()),
            new Run(new Text(" " + text) { Space = SpaceProcessingModeValues.Preserve }));
        ep.Endnotes.Append(new Endnote(np) { Id = id });
        Changed();
    }

    /// <summary>Sposta il cursore sul richiamo di nota (a piè di pagina o di chiusura) successivo; riparte dall'inizio in fondo al documento. False se non ci sono note.</summary>
    public bool NextNote()
    {
        var refs = Doc.Body.Descendants<Run>().Where(r => r.GetFirstChild<FootnoteReference>() != null || r.GetFirstChild<EndnoteReference>() != null).ToList();
        if (refs.Count == 0) return false;
        var ps = Paragraphs();
        var cand = refs.Select(r => (r, p: r.Ancestors<Paragraph>().FirstOrDefault())).Where(x => x.p != null && ps.Contains(x.p!)).Select(x => (x.r, p: x.p!, off: OffsetBefore(x.p!, x.r))).ToList();
        if (cand.Count == 0) return false;
        var next = cand.FirstOrDefault(c => Cmp(new Pos(c.p, c.off), Caret) > 0);
        if (next.p == null) next = cand[0];
        Set(new Pos(next.p, next.off));
        return true;
    }

    // ---------- forme ----------
    /// <summary>Inserisce una forma predefinita (una di Shapes.Gallery) ancorata al paragrafo del cursore, con testo sopra e sotto. Dimensioni in punti; fill/line RRGGBB o null.</summary>
    public void InsertShape(string prst, double widthPt, double heightPt, string? fill, string? line)
    {
        if (!PowerNotes.Core.Layout.Shapes.IsKnown(prst)) throw new ArgumentException("Forma non riconosciuta.", nameof(prst));
        if (double.IsNaN(widthPt) || double.IsNaN(heightPt) || widthPt < 4 || heightPt < 4 || widthPt > 1584 || heightPt > 1584) throw new ArgumentOutOfRangeException(nameof(widthPt), "Dimensioni della forma da 4 a 1584 pt.");
        foreach (var c in new[] { fill, line }) if (c != null && !Regex.IsMatch(c, "^[0-9A-Fa-f]{6}$")) throw new ArgumentException("Colore non valido: usa 6 cifre esadecimali.");
        if (Caret.P.Ancestors<Body>().Count() == 0 || Caret.P.Ancestors<Footnote>().Any() || Caret.P.Ancestors<Endnote>().Any()) throw new InvalidOperationException("Le forme si inseriscono solo nel corpo del documento.");
        Begin();
        if (HasSelection) DeleteSelectionCore();
        long cx = (long)(widthPt * 12700), cy = (long)(heightPt * 12700);
        uint id = (uint)(Doc.Body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().Select(d => d.Id?.Value ?? 0).DefaultIfEmpty(0u).Max() + 1);
        string fillXml = fill == null ? "<a:noFill/>" : $"<a:solidFill><a:srgbClr val=\"{fill.ToUpperInvariant()}\"/></a:solidFill>";
        string lineXml = line == null ? "<a:ln><a:noFill/></a:ln>" : $"<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"{line.ToUpperInvariant()}\"/></a:solidFill></a:ln>";
        var xml = "<w:drawing xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
            "<wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251659264\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">" +
            "<wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV>" +
            $"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:wrapTopAndBottom/><wp:docPr id=\"{id}\" name=\"Forma {id}\"/><wp:cNvGraphicFramePr/>" +
            $"<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"{prst}\"><a:avLst/></a:prstGeom>{fillXml}{lineXml}</wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>";
        XmlOps.PlaceRun(Caret.P, Caret.Off, new Run(new Drawing(xml)));
        Changed();
    }

    // ---------- numeri di riga ----------
    public enum LineNumMode { None, Continuous, RestartPage, RestartSection }

    public LineNumMode GetLineNumbers()
    {
        var ln = SectionFor(Caret.P).GetFirstChild<LineNumberType>();
        if (ln == null) return LineNumMode.None;
        var r = ln.Restart?.Value;
        return r == LineNumberRestartValues.Continuous ? LineNumMode.Continuous : r == LineNumberRestartValues.NewSection ? LineNumMode.RestartSection : LineNumMode.RestartPage;
    }

    /// <summary>Numeri di riga per tutto il documento (tutte le sezioni): spenti, continui, da riavviare a ogni pagina o a ogni sezione.</summary>
    public void SetLineNumbers(LineNumMode mode)
    {
        Begin();
        foreach (var sp in AllSections())
        {
            sp.GetFirstChild<LineNumberType>()?.Remove();
            if (mode == LineNumMode.None) continue;
            var ln = new LineNumberType { CountBy = 1, Restart = mode == LineNumMode.Continuous ? LineNumberRestartValues.Continuous : mode == LineNumMode.RestartSection ? LineNumberRestartValues.NewSection : LineNumberRestartValues.NewPage };
            // ordine schema: dopo pgMar/paperSrc/pgBorders, prima di pgNumType/cols
            var after = sp.Elements().LastOrDefault(e => e is PageMargin or PaperSource or PageBorders) ?? sp.Elements().LastOrDefault(e => e is PageSize or HeaderReference or FooterReference or FootnoteProperties or EndnoteProperties or SectionType);
            if (after != null) after.InsertAfterSelf(ln); else sp.PrependChild(ln);
        }
        Changed();
    }

    // ---------- riquadro di spostamento ----------
    /// <summary>Titoli del corpo (paragrafi con livello struttura 1-9, non vuoti), in ordine di documento.</summary>
    public List<(Paragraph P, int Level, string Text)> Headings()
    {
        var res = new List<(Paragraph, int, string)>();
        foreach (var p in Doc.Paragraphs())
        {
            var lvl = Styles.ResolvePara(p, null, out _).OutlineLevel;
            if (lvl is < 0 or > 8) continue;
            var t = ParaText.Text(p).Replace("\n", " ").Replace("\t", " ").Trim();
            if (t.Length > 0) res.Add((p, lvl, t));
        }
        return res;
    }
}
