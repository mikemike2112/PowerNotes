using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Layout;
using PowerNotes.Core.Model;

namespace PowerNotes.Core.Edit;

public readonly record struct Pos(Paragraph P, int Off);

public enum PageNumFormat { None, Number, PageOfTotal }

/// <summary>Controller dell'editor: caret, selezione, comandi di modifica. Indipendente da WPF, quindi testabile.</summary>
public sealed class EditorSession
{
    public WordDoc Doc { get; }
    public Pos Caret { get; private set; }
    public Pos Anchor { get; private set; }
    public RunProperties? Pending { get; private set; }
    double _goalX = -1;
    bool _preferPrev;
    LayoutResult? _layout; int _layoutVer = -1;
    List<Paragraph>? _paras; int _parasVer = -1;
    StyleResolver? _st; int _stVer = -1;

    public EditorSession(WordDoc doc)
    {
        Doc = doc;
        var first = Paragraphs().FirstOrDefault() ?? EnsureParagraph();
        Caret = Anchor = new Pos(first, 0);
    }

    Paragraph EnsureParagraph()
    {
        var p = new Paragraph();
        Doc.Body.InsertAt(p, 0);
        Doc.Touched();
        return p;
    }

    public LayoutResult Layout
    {
        get
        {
            if (_layout == null || _layoutVer != Doc.Version) { _layout = Doc.Layout(); _layoutVer = Doc.Version; }
            return _layout;
        }
    }

    OpenXmlElement? _storyRoot; Pos _bodyCaret;
    /// <summary>True mentre si modifica un'intestazione o un piè di pagina direttamente sul foglio.</summary>
    public bool EditingHF => _storyRoot != null;
    public bool EditingHeader => _storyRoot is Header;

    public List<Paragraph> Paragraphs()
    {
        if (_paras == null || _parasVer != Doc.Version || _parasRoot != _storyRoot)
        {
            _paras = _storyRoot == null ? Doc.Paragraphs() : _storyRoot.Descendants<Paragraph>().Where(p => !p.Ancestors<Run>().Any()).ToList();
            _parasVer = Doc.Version; _parasRoot = _storyRoot;
        }
        return _paras;
    }
    OpenXmlElement? _parasRoot;

    /// <summary>Doppio clic nella zona dell'intestazione/piè di pagina: crea la parte se manca ed entra in modifica sul foglio.</summary>
    public void EnterHeaderFooter(bool header)
    {
        if (EditingHF) ExitHeaderFooter();
        var main = Doc.Pkg.MainDocumentPart!;
        var sp = SectionFor(Caret.P);
        var part = FindHF(sp, header);
        if (part == null) { SetHeaderFooter(header, "", Align.Left, PageNumFormat.None); part = FindHF(sp, header); }
        OpenXmlElement? root = header ? (part as HeaderPart)?.Header : (OpenXmlElement?)(part as FooterPart)?.Footer;
        if (root == null) return;
        _bodyCaret = Caret;
        _storyRoot = root;
        var first = root.Descendants<Paragraph>().First();
        Pending = null; Caret = Anchor = new Pos(first, 0);
        _paras = null; Doc.Touched();
    }

    public void ExitHeaderFooter()
    {
        if (!EditingHF) return;
        _storyRoot = null; _paras = null;
        var ps = Paragraphs();
        var back = ps.Contains(_bodyCaret.P) ? _bodyCaret : new Pos(ps[0], 0);
        Caret = Anchor = back; Pending = null; Doc.Touched();
    }

    public StyleResolver Styles
    {
        get { if (_st == null || _stVer != Doc.Version) { _st = new StyleResolver(Doc.Pkg); _stVer = Doc.Version; } return _st; }
    }

    int Idx(Paragraph p) => Math.Max(0, Paragraphs().IndexOf(p));

    int Cmp(Pos a, Pos b) { int c = Idx(a.P).CompareTo(Idx(b.P)); return c != 0 ? c : a.Off.CompareTo(b.Off); }

    public bool HasSelection => Caret.P != Anchor.P || Caret.Off != Anchor.Off;
    public (Pos Start, Pos End) Range => Cmp(Caret, Anchor) <= 0 ? (Caret, Anchor) : (Anchor, Caret);

    // ---------- modifica: infrastruttura ----------
    void Begin(string? key = null) => Doc.BeginEdit(Idx(Caret.P), Caret.Off, Idx(Anchor.P), Anchor.Off, key);

    void Set(Pos p, bool keepAnchor = false, bool keepGoal = false)
    {
        Caret = p;
        if (!keepAnchor) Anchor = p;
        if (!keepGoal) _goalX = -1;
        _preferPrev = false;
        Pending = null;
    }

    void Changed() => Doc.Touched();

    public void SetCaret(Paragraph p, int off, bool extend = false) { Set(new Pos(p, Math.Clamp(off, 0, ParaText.Length(p))), extend); }

    public void SetCaretAt(int page, double x, double y, bool extend = false)
    {
        var hit = Layout.HitTest(page, x, y, EditingHF);
        if (hit == null) return;
        Set(new Pos(hit.Value.Para, hit.Value.Offset), extend);
        _preferPrev = hit.Value.Offset == hit.Value.Line.End && !hit.Value.Line.LastOfPara;
    }

    public void SelectAll()
    {
        var ps = Paragraphs();
        Anchor = new Pos(ps[0], 0); Caret = new Pos(ps[^1], ParaText.Length(ps[^1]));
    }

    public void SelectWord()
    {
        var t = ParaText.Text(Caret.P); int i = Math.Min(Caret.Off, t.Length);
        int a = i, b = i;
        while (a > 0 && char.IsLetterOrDigit(t[a - 1])) a--;
        while (b < t.Length && char.IsLetterOrDigit(t[b])) b++;
        Anchor = new Pos(Caret.P, a); Caret = new Pos(Caret.P, b);
    }

    public string SelectedText()
    {
        var (s, e) = Range;
        if (s.P == e.P) return ParaText.Text(s.P)[s.Off..Math.Min(e.Off, ParaText.Text(s.P).Length)];
        var ps = Paragraphs(); int i1 = Idx(s.P), i2 = Idx(e.P);
        var sb = new System.Text.StringBuilder();
        for (int i = i1; i <= i2; i++)
        {
            var t = ParaText.Text(ps[i]);
            int a = i == i1 ? s.Off : 0, b = i == i2 ? Math.Min(e.Off, t.Length) : t.Length;
            sb.Append(t[Math.Min(a, t.Length)..b]);
            if (i < i2) sb.Append('\n');
        }
        return sb.ToString();
    }

    // ---------- digitazione ----------
    public void Type(string text)
    {
        if (text.Length == 0) return;
        bool plain = !text.Contains('\n') && !text.Contains('\t') && !text.Contains('\r');
        Begin(plain && !HasSelection ? "type" : null);
        if (HasSelection) DeleteSelectionCore();
        foreach (var part in System.Text.RegularExpressions.Regex.Split(text.Replace("\r\n", "\n").Replace('\r', '\n'), "(\n|\t)"))
        {
            if (part == "\n") EnterCore();
            else if (part == "\t") { XmlOps.InsertTab(Caret.P, Caret.Off); Set(new Pos(Caret.P, Caret.Off + 1)); }
            else if (part.Length > 0)
            {
                XmlOps.InsertText(Caret.P, Caret.Off, part, Pending);
                Set(new Pos(Caret.P, Caret.Off + part.Length));
            }
        }
        Changed();
    }

    public void SoftBreak()
    {
        Begin();
        if (HasSelection) DeleteSelectionCore();
        XmlOps.InsertBreak(Caret.P, Caret.Off, null);
        Set(new Pos(Caret.P, Caret.Off + 1)); Changed();
    }

    public void Enter() { Begin(); if (HasSelection) DeleteSelectionCore(); EnterCore(); Changed(); }

    void EnterCore()
    {
        var p = Caret.P;
        var pPr = p.ParagraphProperties;
        if (ParaText.Length(p) == 0 && pPr?.NumberingProperties != null)
        {
            // Invio su elemento vuoto di un elenco: esce dall'elenco
            pPr.NumberingProperties = null;
            if (IsListPara(pPr.ParagraphStyleId?.Val?.Value)) pPr.ParagraphStyleId = null;
            if (pPr.Indentation != null) pPr.Indentation = null;
            return;
        }
        int len = ParaText.Length(p);
        var np = XmlOps.Split(p, Caret.Off);
        if (Caret.Off >= len)
        {
            var next = Styles.Find(pPr?.ParagraphStyleId?.Val?.Value)?.NextParagraphStyle?.Val?.Value;
            if (next != null) { var npr = XmlOps.PPr(np); npr.ParagraphStyleId = next == "Normal" ? null : new ParagraphStyleId { Val = next }; }
        }
        Set(new Pos(np, 0));
    }

    bool IsListPara(string? id) => string.Equals(Styles.Find(id)?.StyleName?.Val?.Value, "List Paragraph", StringComparison.OrdinalIgnoreCase);

    public void Backspace()
    {
        if (HasSelection) { Begin(); DeleteSelectionCore(); Changed(); return; }
        var p = Caret.P;
        if (Caret.Off > 0) { Begin("bs"); XmlOps.DeleteRange(p, Caret.Off - 1, Caret.Off); Set(new Pos(p, Caret.Off - 1)); Changed(); return; }
        var pPr = p.ParagraphProperties;
        if (pPr?.NumberingProperties != null || (pPr?.Indentation?.Left?.Value ?? "0") != "0" && pPr?.Indentation != null)
        {
            Begin();
            if (pPr.NumberingProperties != null) { pPr.NumberingProperties = null; if (IsListPara(pPr.ParagraphStyleId?.Val?.Value)) pPr.ParagraphStyleId = null; }
            pPr.Indentation = null;
            Changed(); return;
        }
        if (p.PreviousSibling<OpenXmlElement>() is Paragraph prev && prev.ParagraphProperties?.SectionProperties == null)
        {
            Begin();
            int l = ParaText.Length(prev);
            XmlOps.Merge(prev, p);
            Set(new Pos(prev, l)); Changed();
        }
    }

    public void DeleteForward()
    {
        if (HasSelection) { Begin(); DeleteSelectionCore(); Changed(); return; }
        var p = Caret.P; int len = ParaText.Length(p);
        if (Caret.Off < len) { Begin("del"); XmlOps.DeleteRange(p, Caret.Off, Caret.Off + 1); Changed(); return; }
        if (p.NextSibling<OpenXmlElement>() is Paragraph next && p.ParagraphProperties?.SectionProperties == null)
        {
            Begin(); XmlOps.Merge(p, next); Changed();
        }
    }

    void DeleteSelectionCore()
    {
        var (s, e) = Range;
        if (s.P == e.P) XmlOps.DeleteRange(s.P, s.Off, e.Off);
        else if (s.P.Parent == e.P.Parent)
        {
            XmlOps.DeleteRange(s.P, s.Off, ParaText.Length(s.P));
            XmlOps.DeleteRange(e.P, 0, e.Off);
            foreach (var mid in s.P.ElementsAfter().TakeWhile(x => x != e.P).ToList()) mid.Remove();
            XmlOps.Merge(s.P, e.P);
        }
        else
        {
            var ps = Paragraphs(); int i1 = Idx(s.P), i2 = Idx(e.P);
            for (int i = i1; i <= i2; i++)
            {
                int a = i == i1 ? s.Off : 0, b = i == i2 ? e.Off : ParaText.Length(ps[i]);
                XmlOps.DeleteRange(ps[i], a, b);
            }
        }
        Set(new Pos(s.P, s.Off));
    }

    public void DeleteSelection() { if (!HasSelection) return; Begin(); DeleteSelectionCore(); Changed(); }

    public string Cut() { var t = SelectedText(); DeleteSelection(); return t; }

    // ---------- appunti con formattazione ----------
    /// <summary>Frammento OOXML (paragrafi ritagliati sulla selezione, con pPr e rPr) per l'incolla formattato. Null senza selezione.</summary>
    public string? CopyFragment()
    {
        if (!HasSelection) return null;
        var (s, e) = Range; var ps = Paragraphs(); int i1 = Idx(s.P), i2 = Idx(e.P);
        var body = new Body();
        for (int i = i1; i <= i2; i++)
        {
            var c = (Paragraph)ps[i].CloneNode(true);
            c.ParagraphProperties?.SectionProperties?.Remove();
            int len = ParaText.Length(c);
            int a = i == i1 ? s.Off : 0, b = i == i2 ? e.Off : len;
            XmlOps.DeleteRange(c, b, len); XmlOps.DeleteRange(c, 0, a);
            body.Append(c);
        }
        return body.OuterXml;
    }

    /// <summary>Incolla un frammento prodotto da CopyFragment. sameDocument=false: le immagini (relazioni di un altro documento) vengono omesse.</summary>
    public void PasteFragment(string xml, bool sameDocument)
    {
        Body frag;
        try { frag = new Body(xml); }
        catch (Exception ex) when (ex is System.Xml.XmlException or ArgumentException or InvalidOperationException or FormatException) { throw new InvalidOperationException("Appunti non validi.", ex); }
        var paras = frag.Elements<Paragraph>().Select(p => (Paragraph)p.CloneNode(true)).ToList();
        if (paras.Count == 0) return;
        if (!sameDocument) foreach (var p in paras) foreach (var d in p.Descendants<Drawing>().ToList()) { var r = d.Ancestors<Run>().FirstOrDefault(); (r ?? (OpenXmlElement)d).Remove(); }
        Begin();
        if (HasSelection) DeleteSelectionCore();
        var cur = Caret.P; int off = Caret.Off;
        var tail = XmlOps.Split(cur, off);
        IEnumerable<OpenXmlElement> Content(Paragraph p) => p.ChildElements.Where(c => c is not ParagraphProperties).ToList();
        foreach (var c in Content(paras[0])) { c.Remove(); cur.Append(c); }
        Pos end;
        if (paras.Count == 1) { end = new Pos(cur, ParaText.Length(cur)); XmlOps.Merge(cur, tail); }
        else
        {
            var prev = cur;
            for (int i = 1; i < paras.Count - 1; i++) { prev.InsertAfterSelf(paras[i]); prev = paras[i]; }
            var last = paras[^1]; prev.InsertAfterSelf(last);
            end = new Pos(last, ParaText.Length(last));
            XmlOps.Merge(last, tail);
        }
        Set(end); Changed();
    }

    // ---------- movimento ----------
    public void MoveHorizontal(int dir, bool extend, bool word = false)
    {
        if (HasSelection && !extend) { var (s, e) = Range; Set(dir < 0 ? s : e); return; }
        var p = Caret.P; int off = Caret.Off;
        var ps = Paragraphs();
        if (word)
        {
            var t = ParaText.Text(p);
            if (dir < 0)
            {
                if (off == 0) { MoveHorizontal(-1, extend); return; }
                while (off > 0 && !char.IsLetterOrDigit(t[off - 1])) off--;
                while (off > 0 && char.IsLetterOrDigit(t[off - 1])) off--;
            }
            else
            {
                if (off >= t.Length) { MoveHorizontal(1, extend); return; }
                while (off < t.Length && char.IsLetterOrDigit(t[off])) off++;
                while (off < t.Length && !char.IsLetterOrDigit(t[off])) off++;
            }
            Set(new Pos(p, off), extend); return;
        }
        if (dir < 0)
        {
            if (off > 0) Set(new Pos(p, off - 1), extend);
            else { int i = Idx(p); if (i > 0) Set(new Pos(ps[i - 1], ParaText.Length(ps[i - 1])), extend); }
        }
        else
        {
            if (off < ParaText.Length(p)) Set(new Pos(p, off + 1), extend);
            else { int i = Idx(p); if (i < ps.Count - 1) Set(new Pos(ps[i + 1], 0), extend); }
        }
    }

    public void MoveVertical(int dir, bool extend, double pageStep = 0)
    {
        var L = Layout;
        var line = L.Locate(Caret.P, Caret.Off, _preferPrev);
        if (line == null) return;
        double gx = _goalX >= 0 ? _goalX : line.XAt(Caret.Off);
        LineBox? target = line;
        int steps = pageStep > 0 ? Math.Max(1, (int)(pageStep / Math.Max(1, line.Height))) : 1;
        for (int s = 0; s < steps; s++)
        {
            var cur = target!;
            LineBox? best = null; double bestKey = double.MaxValue, bestDx = double.MaxValue;
            for (int pi = Math.Max(0, cur.PageIndex - 1); pi <= Math.Min(L.Pages.Count - 1, cur.PageIndex + 1); pi++)
                foreach (var l in L.Pages[pi].Lines)
                {
                    double rel = (pi - cur.PageIndex) * 100000.0 + l.Y;
                    double curY = cur.Y;
                    bool ok = dir > 0 ? rel > curY + 0.5 : rel < curY - 0.5;
                    if (!ok) continue;
                    double key = Math.Abs(rel - curY);
                    double dx = gx < l.Left ? l.Left - gx : gx > l.Right ? gx - l.Right : 0;
                    if (key < bestKey - 0.5 || (Math.Abs(key - bestKey) <= 0.5 && dx < bestDx)) { best = l; bestKey = key; bestDx = dx; }
                }
            if (best == null) break;
            target = best;
        }
        if (target == line)
        {
            // prima/ultima riga: va a inizio/fine documento
            var ps = Paragraphs();
            if (dir < 0) Set(new Pos(ps[0], 0), extend, true); else Set(new Pos(ps[^1], ParaText.Length(ps[^1])), extend, true);
            return;
        }
        int off = target!.OffsetAt(gx);
        Set(new Pos(target.Para, off), extend, true);
        _goalX = gx;
        _preferPrev = off == target.End && !target.LastOfPara;
    }

    public void MoveLineEdge(bool end, bool extend, bool doc = false)
    {
        var ps = Paragraphs();
        if (doc) { Set(end ? new Pos(ps[^1], ParaText.Length(ps[^1])) : new Pos(ps[0], 0), extend); return; }
        var line = Layout.Locate(Caret.P, Caret.Off, _preferPrev);
        if (line == null) return;
        Set(new Pos(line.Para, end ? line.End : line.Start), extend);
        _preferPrev = end && !line.LastOfPara;
    }

    // ---------- undo / redo ----------
    public bool Undo() => Restore(Doc.Undo(Idx(Caret.P), Caret.Off, Idx(Anchor.P), Anchor.Off));
    public bool Redo() => Restore(Doc.Redo(Idx(Caret.P), Caret.Off, Idx(Anchor.P), Anchor.Off));

    bool Restore(WordDoc.Snap? s)
    {
        if (s == null) return false;
        _storyRoot = null; _paras = null;
        var ps = Paragraphs();
        if (ps.Count == 0) { EnsureParagraph(); ps = Paragraphs(); }
        Pos P(int i, int o) { var p = ps[Math.Clamp(i, 0, ps.Count - 1)]; return new Pos(p, Math.Clamp(o, 0, ParaText.Length(p))); }
        Anchor = P(s.AnchorPara, s.AnchorOff); Caret = P(s.CaretPara, s.CaretOff);
        Pending = null;
        return true;
    }

    // ---------- formattazione carattere ----------
    IEnumerable<Paragraph> SelectedParagraphs()
    {
        var (s, e) = Range; var ps = Paragraphs();
        for (int i = Idx(s.P); i <= Idx(e.P); i++) yield return ps[i];
    }

    public void ApplyRun(Action<RunProperties> f)
    {
        Begin();
        if (!HasSelection)
        {
            var leaves = ParaText.Leaves(Caret.P);
            var src = leaves.LastOrDefault(l => l.Start < Caret.Off)?.Run.RunProperties ?? leaves.FirstOrDefault()?.Run.RunProperties;
            var rp = src != null ? (RunProperties)src.CloneNode(true) : new RunProperties();
            f(rp); Pending = rp;
            if (ParaText.Length(Caret.P) == 0)
            {
                // paragrafo vuoto: la formattazione va anche sul segno di paragrafo
                var pm = XmlOps.PPr(Caret.P).ParagraphMarkRunProperties ?? (XmlOps.PPr(Caret.P).ParagraphMarkRunProperties = new ParagraphMarkRunProperties());
                var tmp = new RunProperties(); f(tmp);
                foreach (var c in tmp.ChildElements.ToList()) { foreach (var old in pm.ChildElements.Where(x => x.GetType() == c.GetType()).ToList()) old.Remove(); pm.Append(c.CloneNode(true)); }
            }
            Changed(); return;
        }
        var (s, e) = Range;
        foreach (var p in SelectedParagraphs())
        {
            int a = p == s.P ? s.Off : 0, b = p == e.P ? e.Off : ParaText.Length(p);
            XmlOps.Format(p, a, b, f);
        }
        // l'offset dei confini non cambia: le posizioni restano valide
        Changed();
    }

    /// <summary>Formattazione effettiva al caret (o alla selezione: vale solo se uniforme).</summary>
    public RunFmt FormatAtCaret()
    {
        var p = Caret.P;
        Styles.ResolvePara(p, null, out var mark, out var baseRun);
        var leaves = ParaText.Leaves(p);
        var leaf = leaves.LastOrDefault(l => l.Start < Caret.Off) ?? leaves.FirstOrDefault();
        var fmt = leaf == null ? mark : Styles.ResolveRun(baseRun, leaf.Run);
        if (Pending != null) { fmt = fmt.Clone(); Styles.ApplyRPr(fmt, Pending); }
        return fmt;
    }

    public ParaFmt ParaFormatAtCaret() => Styles.ResolvePara(Caret.P, null, out _);

    /// <summary>True se tutto il testo selezionato soddisfa il predicato (oppure, senza selezione, il carattere al caret).</summary>
    public bool AllRuns(Func<RunFmt, bool> pred)
    {
        if (!HasSelection) return pred(FormatAtCaret());
        var (s, e) = Range; bool any = false;
        foreach (var p in SelectedParagraphs())
        {
            Styles.ResolvePara(p, null, out _, out var mark);
            int a = p == s.P ? s.Off : 0, b = p == e.P ? e.Off : ParaText.Length(p);
            foreach (var l in ParaText.Leaves(p).Where(l => l.End > a && l.Start < b))
            {
                any = true;
                if (!pred(Styles.ResolveRun(mark, l.Run))) return false;
            }
        }
        return any;
    }

    public void ToggleBold() { bool on = !AllRuns(f => f.Bold); ApplyRun(r => r.Bold = new Bold { Val = on }); }
    public void ToggleItalic() { bool on = !AllRuns(f => f.Italic); ApplyRun(r => r.Italic = new Italic { Val = on }); }
    public void ToggleUnderline() { bool on = !AllRuns(f => f.Underline); ApplyRun(r => r.Underline = new Underline { Val = on ? UnderlineValues.Single : UnderlineValues.None }); }
    public void ToggleStrike() { bool on = !AllRuns(f => f.Strike); ApplyRun(r => r.Strike = new Strike { Val = on }); }
    static readonly double[] SizeSteps = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

    /// <summary>Ingrandisci/Riduci carattere: passa al valore successivo/precedente dell'elenco standard.</summary>
    public void StepSize(bool grow)
    {
        double cur = FormatAtCaret().Size;
        double next = grow ? SizeSteps.FirstOrDefault(x => x > cur + 0.01) : SizeSteps.LastOrDefault(x => x < cur - 0.01);
        if (next == 0) next = cur;
        SetSize(next);
    }

    public void SetCaps(bool on, bool small) => ApplyRun(r => { r.Caps = small ? null : on ? new Caps() : null; r.SmallCaps = small && on ? new SmallCaps() : null; if (!small && !on) r.Caps = null; });
    public void SetFont(string name) => ApplyRun(r => r.RunFonts = new RunFonts { Ascii = name, HighAnsi = name, EastAsia = name, ComplexScript = name });
    public void SetSize(double pt) => ApplyRun(r => { var v = ((int)Math.Round(pt * 2)).ToString(); r.FontSize = new FontSize { Val = v }; r.FontSizeComplexScript = new FontSizeComplexScript { Val = v }; });
    public void SetColor(string? hex) => ApplyRun(r => r.Color = hex == null ? new Color { Val = "auto" } : new Color { Val = hex });
    public void SetHighlight(HighlightColorValues? h) => ApplyRun(r => r.Highlight = h == null ? null : new Highlight { Val = h });
    public void SetVertAlign(int v) => ApplyRun(r => r.VerticalTextAlignment = v == 0 ? null : new VerticalTextAlignment { Val = v > 0 ? VerticalPositionValues.Superscript : VerticalPositionValues.Subscript });

    // ---------- formattazione paragrafo ----------
    void ApplyPara(Action<ParagraphProperties> f)
    {
        Begin();
        foreach (var p in SelectedParagraphs()) f(XmlOps.PPr(p));
        Changed();
    }

    public void SetAlignment(Align a) => ApplyPara(pp => pp.Justification = new Justification { Val = a switch { Align.Center => JustificationValues.Center, Align.Right => JustificationValues.Right, Align.Justify => JustificationValues.Both, _ => JustificationValues.Left } });

    /// <summary>Interlinea: rule auto (line=240esimi di riga), exact o atLeast (twips).</summary>
    public void SetLineSpacing(double multiple) => ApplyPara(pp =>
    {
        var sp = pp.SpacingBetweenLines ?? (pp.SpacingBetweenLines = new SpacingBetweenLines());
        sp.Line = ((int)Math.Round(multiple * 240)).ToString(); sp.LineRule = LineSpacingRuleValues.Auto;
    });

    public void SetSpacing(double? beforePt, double? afterPt) => ApplyPara(pp =>
    {
        var sp = pp.SpacingBetweenLines ?? (pp.SpacingBetweenLines = new SpacingBetweenLines());
        if (beforePt != null) sp.Before = ((int)Math.Round(beforePt.Value * 20)).ToString();
        if (afterPt != null) sp.After = ((int)Math.Round(afterPt.Value * 20)).ToString();
    });

    /// <summary>Rientri in punti. first positivo = prima riga, negativo = sporgente.</summary>
    public void SetIndent(double? left, double? right, double? first) => ApplyPara(pp =>
    {
        var ind = pp.Indentation ?? (pp.Indentation = new Indentation());
        if (left != null) ind.Left = ((int)Math.Round(left.Value * 20)).ToString();
        if (right != null) ind.Right = ((int)Math.Round(right.Value * 20)).ToString();
        if (first != null)
        {
            ind.FirstLine = null; ind.Hanging = null;
            if (first > 0) ind.FirstLine = ((int)Math.Round(first.Value * 20)).ToString();
            else if (first < 0) ind.Hanging = ((int)Math.Round(-first.Value * 20)).ToString();
        }
    });

    public void ChangeIndent(double deltaPt)
    {
        Begin();
        foreach (var p in SelectedParagraphs())
        {
            var f = Styles.ResolvePara(p, null, out _);
            var pp = XmlOps.PPr(p);
            if (pp.NumberingProperties != null)
            {
                int lvl = Math.Clamp((pp.NumberingProperties.NumberingLevelReference?.Val?.Value ?? f.Ilvl ?? 0) + (deltaPt > 0 ? 1 : -1), 0, 8);
                pp.NumberingProperties.NumberingLevelReference = new NumberingLevelReference { Val = lvl };
            }
            else { var ind = pp.Indentation ?? (pp.Indentation = new Indentation()); ind.Left = ((int)Math.Round(Math.Max(0, f.IndLeft + deltaPt) * 20)).ToString(); }
        }
        Changed();
    }

    public void SetStyle(string styleId)
    {
        var real = BuiltIns.EnsureStyle(Doc.Pkg, styleId);
        bool isDefault = Styles.Find(real)?.Default?.Value == true;
        ApplyPara(pp =>
        {
            pp.ParagraphStyleId = isDefault ? null : new ParagraphStyleId { Val = real };
            pp.Justification = null; pp.SpacingBetweenLines = null; pp.Indentation = null; pp.KeepNext = null; pp.KeepLines = null;
        });
    }

    public void ToggleList(string kind)
    {
        int numId = BuiltIns.EnsureList(Doc.Pkg, kind);
        var ps = SelectedParagraphs().ToList();
        bool all = ps.All(p => BuiltIns.KindOf(Doc.Pkg, Styles.ResolvePara(p, null, out _).NumId) == kind);
        bool fresh = !all && kind != "bullet" && !(ps[0].PreviousSibling<OpenXmlElement>() is Paragraph pv && BuiltIns.KindOf(Doc.Pkg, pv.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value) == kind);
        if (fresh) numId = BuiltIns.NewInstance(Doc.Pkg, numId);
        var lpId = BuiltIns.EnsureStyle(Doc.Pkg, "ListParagraph");
        ApplyPara(pp =>
        {
            if (all) { pp.NumberingProperties = null; if (IsListPara(pp.ParagraphStyleId?.Val?.Value)) pp.ParagraphStyleId = null; pp.Indentation = null; }
            else
            {
                pp.NumberingProperties = new NumberingProperties(new NumberingLevelReference { Val = pp.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0 }, new NumberingId { Val = numId });
                if (pp.ParagraphStyleId == null) pp.ParagraphStyleId = new ParagraphStyleId { Val = lpId };
            }
        });
    }

    // ---------- inserimenti ----------
    public void InsertPageBreak()
    {
        Begin();
        if (HasSelection) DeleteSelectionCore();
        XmlOps.InsertBreak(Caret.P, Caret.Off, BreakValues.Page);
        var np = XmlOps.Split(Caret.P, Caret.Off + 1);
        Set(new Pos(np, 0)); Changed();
    }

    public void InsertTable(int rows, int cols)
    {
        if (rows < 1 || cols < 1 || rows > 500 || cols > 63) throw new ArgumentOutOfRangeException(nameof(rows), "Dimensioni tabella non valide.");
        Begin();
        if (HasSelection) DeleteSelectionCore();
        var gridId = BuiltIns.EnsureStyle(Doc.Pkg, "TableGrid");
        var p = Caret.P;
        var sec = SectionFor(p);
        double w = ((sec.GetFirstChild<PageSize>()?.Width?.Value ?? 11906) - (sec.GetFirstChild<PageMargin>()?.Left?.Value ?? 1440) - (sec.GetFirstChild<PageMargin>()?.Right?.Value ?? 1440));
        if (p.Parent is TableCell) w = 4000;
        int cw = (int)(w / cols);
        var sb = new System.Text.StringBuilder($"<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tblPr><w:tblStyle w:val=\"{gridId}\"/><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/></w:tblPr><w:tblGrid>");
        for (int c = 0; c < cols; c++) sb.Append($"<w:gridCol w:w=\"{cw}\"/>");
        sb.Append("</w:tblGrid>");
        for (int r = 0; r < rows; r++)
        {
            sb.Append("<w:tr>");
            for (int c = 0; c < cols; c++) sb.Append($"<w:tc><w:tcPr><w:tcW w:w=\"{cw}\" w:type=\"dxa\"/></w:tcPr><w:p/></w:tc>");
            sb.Append("</w:tr>");
        }
        sb.Append("</w:tbl>");
        var tbl = new Table(sb.ToString());
        bool empty = ParaText.Length(p) == 0;
        if (empty) p.InsertBeforeSelf(tbl);
        else
        {
            var tail = XmlOps.Split(p, Caret.Off);
            p.InsertAfterSelf(tbl);   // tail resta dopo la tabella
        }
        var first = tbl.Descendants<Paragraph>().First();
        Set(new Pos(first, 0)); Changed();
    }

    /// <summary>Cella (riga, cella, colonna di griglia iniziale, estensione) dei figli di una riga.</summary>
    static List<(TableCell Cell, int Col, int Span)> Cells(TableRow row)
    {
        var l = new List<(TableCell, int, int)>();
        int col = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
        foreach (var c in row.Elements<TableCell>()) { int sp = Math.Max(1, c.TableCellProperties?.GridSpan?.Val?.Value ?? 1); l.Add((c, col, sp)); col += sp; }
        return l;
    }

    (TableCell cell, TableRow row, Table tbl)? CellOf(Paragraph p)
    {
        var c = p.Ancestors<TableCell>().FirstOrDefault();
        return c == null ? null : (c, (TableRow)c.Parent!, (Table)c.Parent!.Parent!);
    }

    public bool InTable => CellOf(Caret.P) != null;

    /// <summary>Caret all'inizio di un elemento di elenco (Tab lì cambia livello).</summary>
    public bool AtListStart => !HasSelection && Caret.Off == 0 && Styles.ResolvePara(Caret.P, null, out _).NumId is > 0;

    /// <summary>Tab in tabella: cella successiva (o precedente); nell'ultima cella aggiunge una riga. False se non si è in tabella.</summary>
    public bool TabInTable(bool back)
    {
        var c = CellOf(Caret.P); if (c == null) return false;
        var cells = c.Value.tbl.Elements<TableRow>().SelectMany(r => r.Elements<TableCell>()).ToList();
        int i = cells.IndexOf(c.Value.cell) + (back ? -1 : 1);
        if (i < 0) return true;
        if (i >= cells.Count)
        {
            AddRowBelow();
            cells = c.Value.tbl.Elements<TableRow>().SelectMany(r => r.Elements<TableCell>()).ToList();
        }
        var p = cells[Math.Min(i, cells.Count - 1)].Descendants<Paragraph>().First();
        Set(new Pos(p, 0));
        return true;
    }

    /// <summary>Unisce le celle comprese nella selezione (rettangolo in coordinate di griglia).</summary>
    public bool MergeCells()
    {
        var a = CellOf(Anchor.P); var b = CellOf(Caret.P);
        if (a == null || b == null || a.Value.tbl != b.Value.tbl) return false;
        var tbl = a.Value.tbl; var rows = tbl.Elements<TableRow>().ToList();
        (int r, int c0, int c1) Pos(TableCell cell, TableRow row) { var e = Cells(row).First(x => x.Cell == cell); return (rows.IndexOf(row), e.Col, e.Col + e.Span); }
        var pa = Pos(a.Value.cell, a.Value.row); var pb = Pos(b.Value.cell, b.Value.row);
        int r0 = Math.Min(pa.r, pb.r), r1 = Math.Max(pa.r, pb.r), g0 = Math.Min(pa.c0, pb.c0), g1 = Math.Max(pa.c1, pb.c1);
        if (r0 == r1 && g0 == g1 - 1 || (pa.r == pb.r && pa.c0 == pb.c0)) return false;
        Begin();
        TableCell? top = null;
        for (int r = r0; r <= r1; r++)
        {
            var inRect = Cells(rows[r]).Where(x => x.Col >= g0 && x.Col + x.Span <= g1).ToList();
            if (inRect.Count == 0) continue;
            var keep = inRect[0].Cell;
            int span = inRect.Sum(x => x.Span);
            // sposta il contenuto non vuoto nella cella che resta
            foreach (var x in inRect.Skip(1))
            {
                foreach (var para in x.Cell.Elements<Paragraph>().Where(pp => ParaText.Length(pp) > 0).ToList()) { para.Remove(); keep.Append(para); }
                x.Cell.Remove();
            }
            var tcPr = keep.TableCellProperties ?? keep.PrependChild(new TableCellProperties());
            tcPr.GridSpan = span > 1 ? new GridSpan { Val = span } : null;
            tcPr.TableCellWidth = new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = (tcPr.TableCellWidth?.Width?.Value is string ow && int.TryParse(ow, out var ov) ? ov : 0 + 0).ToString() };
            if (r0 != r1) tcPr.VerticalMerge = r == r0 ? new VerticalMerge { Val = MergedCellValues.Restart } : new VerticalMerge();
            if (r == r0) top = keep;
            // cella vuota di riserva: ogni cella deve contenere un paragrafo
            if (!keep.Elements<Paragraph>().Any()) keep.Append(new Paragraph());
            // rimuovi paragrafi vuoti in eccesso se c'è contenuto
            var ps = keep.Elements<Paragraph>().ToList();
            if (ps.Count > 1) foreach (var e in ps.Where(pp => ParaText.Length(pp) == 0).Take(ps.Count - 1).ToList()) if (keep.Elements<Paragraph>().Count() > 1) e.Remove();
        }
        // larghezza: somma delle colonne di griglia spannate
        var grid = tbl.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Select(x => int.TryParse(x.Width?.Value, out var gw) ? gw : 0).ToList();
        for (int r = r0; r <= r1; r++)
            foreach (var x in Cells(rows[r]).Where(x => x.Col >= g0 && x.Col + x.Span <= g1 && x.Span > 1 || x.Col == g0))
            {
                if (x.Col < g0 || x.Col + x.Span > g1) continue;
                var tcw = x.Cell.TableCellProperties!.TableCellWidth ??= new TableCellWidth();
                tcw.Type = TableWidthUnitValues.Dxa; tcw.Width = grid.Skip(x.Col).Take(x.Span).Sum().ToString();
            }
        if (top != null) Set(new Pos(top.Descendants<Paragraph>().First(), 0));
        Changed();
        return true;
    }

    public bool AddRowBelow()
    {
        var c = CellOf(Caret.P); if (c == null) return false;
        Begin();
        var row = c.Value.row; var nr = new TableRow();
        if (row.TableRowProperties != null) nr.Append(row.TableRowProperties.CloneNode(true));
        foreach (var x in Cells(row))
        {
            var nc = new TableCell();
            if (x.Cell.TableCellProperties != null)
            {
                var pr = (TableCellProperties)x.Cell.TableCellProperties.CloneNode(true); pr.VerticalMerge = null; nc.Append(pr);
            }
            nc.Append(new Paragraph());
            nr.Append(nc);
        }
        row.InsertAfterSelf(nr); Changed(); return true;
    }

    public bool AddColumnRight() { var c = CellOf(Caret.P); if (c == null) return false; var me = Cells(c.Value.row).First(x => x.Cell == c.Value.cell); return AddColumnAt(c.Value.tbl, me.Col + me.Span, me.Col); }
    public bool AddColumnLeft() { var c = CellOf(Caret.P); if (c == null) return false; var me = Cells(c.Value.row).First(x => x.Cell == c.Value.cell); return AddColumnAt(c.Value.tbl, me.Col, me.Col); }

    /// <summary>Nuova colonna di griglia all'indice at: le celle unite che la attraversano si allargano, le altre ricevono una cella vuota.</summary>
    bool AddColumnAt(Table tbl, int at, int widthFrom)
    {
        Begin();
        var grid = tbl.GetFirstChild<TableGrid>()!;
        var cols = grid.Elements<GridColumn>().ToList();
        int w = int.TryParse(cols[Math.Min(widthFrom, cols.Count - 1)].Width?.Value, out var gw) ? gw : 1440;
        var ng = new GridColumn { Width = w.ToString() };
        if (at < cols.Count) cols[at].InsertBeforeSelf(ng); else grid.Append(ng);
        foreach (var row in tbl.Elements<TableRow>())
        {
            var cs = Cells(row);
            var cross = cs.FirstOrDefault(x => x.Col < at && x.Col + x.Span > at);
            if (cross.Cell != null) { var gs = cross.Cell.TableCellProperties!.GridSpan ??= new GridSpan { Val = cross.Span }; gs.Val = cross.Span + 1; continue; }
            var nc = new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = w.ToString() }), new Paragraph());
            var before = cs.FirstOrDefault(x => x.Col >= at);
            if (before.Cell != null) before.Cell.InsertBeforeSelf(nc); else row.Append(nc);
        }
        Changed(); return true;
    }

    public bool AddRowAbove()
    {
        var c = CellOf(Caret.P); if (c == null) return false;
        Begin();
        var row = c.Value.row; var nr = new TableRow();
        if (row.TableRowProperties != null) nr.Append(row.TableRowProperties.CloneNode(true));
        foreach (var x in Cells(row))
        {
            var nc = new TableCell();
            if (x.Cell.TableCellProperties != null) { var pr = (TableCellProperties)x.Cell.TableCellProperties.CloneNode(true); pr.VerticalMerge = null; nc.Append(pr); }
            nc.Append(new Paragraph()); nr.Append(nc);
        }
        row.InsertBeforeSelf(nr); Changed(); return true;
    }

    public bool DeleteRow()
    {
        var c = CellOf(Caret.P); if (c == null) return false;
        Begin();
        var tbl = c.Value.tbl; var row = c.Value.row;
        var next = row.NextSibling<TableRow>() ?? row.PreviousSibling<TableRow>();
        if (next == null) { DeleteTableCore(tbl); Changed(); return true; }
        row.Remove();
        Set(new Pos(next.Descendants<Paragraph>().First(), 0)); Changed(); return true;
    }

    public bool DeleteTable()
    {
        var c = CellOf(Caret.P); if (c == null) return false;
        Begin(); DeleteTableCore(c.Value.tbl); Changed(); return true;
    }

    void DeleteTableCore(Table tbl)
    {
        var target = tbl.NextSibling<Paragraph>() ?? tbl.PreviousSibling<Paragraph>();
        if (target == null) { target = new Paragraph(); tbl.InsertAfterSelf(target); }
        tbl.Remove();
        Set(new Pos(target, 0));
    }

    // ---------- immagini ----------
    public static (int w, int h, PartTypeInfo type)? ImageInfo(byte[] b)
    {
        if (b.Length > 24 && b[0] == 0x89 && b[1] == 'P') return (b[16] << 24 | b[17] << 16 | b[18] << 8 | b[19], b[20] << 24 | b[21] << 16 | b[22] << 8 | b[23], ImagePartType.Png);
        if (b.Length > 10 && b[0] == 'G' && b[1] == 'I' && b[2] == 'F') return (b[6] | b[7] << 8, b[8] | b[9] << 8, ImagePartType.Gif);
        if (b.Length > 26 && b[0] == 'B' && b[1] == 'M') return (BitConverter.ToInt32(b, 18), Math.Abs(BitConverter.ToInt32(b, 22)), ImagePartType.Bmp);
        if (b.Length > 4 && b[0] == 0xFF && b[1] == 0xD8)
        {
            int i = 2;
            while (i + 9 < b.Length)
            {
                if (b[i] != 0xFF) { i++; continue; }
                int m = b[i + 1];
                if (m is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC) return (b[i + 7] << 8 | b[i + 8], b[i + 5] << 8 | b[i + 6], ImagePartType.Jpeg);
                i += 2 + (b[i + 2] << 8 | b[i + 3]);
            }
        }
        return null;
    }

    public void InsertImage(string path)
    {
        byte[] bytes;
        try { bytes = File.ReadAllBytes(path); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { throw new InvalidOperationException("Impossibile leggere l'immagine: " + e.Message, e); }
        var info = ImageInfo(bytes) ?? throw new InvalidOperationException("Formato immagine non supportato (PNG, JPEG, GIF, BMP).");
        if (info.w <= 0 || info.h <= 0 || info.w > 100000 || info.h > 100000) throw new InvalidOperationException("Dimensioni immagine non valide.");
        Begin();
        if (HasSelection) DeleteSelectionCore();
        var main = Doc.Pkg.MainDocumentPart!;
        var ip = main.AddImagePart(info.type);
        using (var ms = new MemoryStream(bytes)) ip.FeedData(ms);
        string rid = main.GetIdOfPart(ip);
        var sec = SectionFor(Caret.P);
        double maxW = ((sec.GetFirstChild<PageSize>()?.Width?.Value ?? 11906) - (sec.GetFirstChild<PageMargin>()?.Left?.Value ?? 1440) - (sec.GetFirstChild<PageMargin>()?.Right?.Value ?? 1440)) / 20.0;
        double wpt = info.w * 0.75, hpt = info.h * 0.75;   // 96 dpi
        if (wpt > maxW) { hpt *= maxW / wpt; wpt = maxW; }
        long cx = (long)(wpt * 12700), cy = (long)(hpt * 12700);
        uint id = (uint)(Doc.Body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().Select(d => d.Id?.Value ?? 0).DefaultIfEmpty(0u).Max() + 1);
        string name = System.Security.SecurityElement.Escape(System.IO.Path.GetFileName(path));
        var xml = $"<w:drawing xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            $"<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"{cx}\" cy=\"{cy}\"/><wp:docPr id=\"{id}\" name=\"Immagine {id}\" descr=\"{name}\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>" +
            $"<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:nvPicPr><pic:cNvPr id=\"{id}\" name=\"{name}\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
            $"<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>";
        var run = new Run(new Drawing(xml));
        XmlOps.PlaceRun(Caret.P, Caret.Off, run);
        Set(new Pos(Caret.P, Caret.Off + 1)); Changed();
    }

    // ---------- impostazione pagina, intestazioni e piè di pagina ----------
    SectionProperties BodySect()
    {
        var sp = Doc.Body.Elements<SectionProperties>().LastOrDefault();
        if (sp == null)
        {
            sp = new SectionProperties(new PageSize { Width = 11906, Height = 16838 }, new PageMargin { Top = 1417, Right = 1134, Bottom = 1134, Left = 1134, Header = 708, Footer = 708, Gutter = 0 });
            Doc.Body.Append(sp);
        }
        return sp;
    }

    SectionProperties SectionFor(Paragraph p)
    {
        OpenXmlElement top = p; while (top.Parent != null && top.Parent != Doc.Body) top = top.Parent;
        foreach (var e in top.ElementsAfter().Prepend(top))
            if (e is Paragraph pp && pp.ParagraphProperties?.SectionProperties is { } sp) return sp;
        return BodySect();
    }

    public IEnumerable<SectionProperties> AllSections() => Doc.Body.Descendants<SectionProperties>().Where(s => s.Parent is Body || s.Parent is ParagraphProperties).Concat(Doc.Body.Elements<SectionProperties>().Any() ? [] : [BodySect()]).Distinct();

    public record PageInfo(double Width, double Height, bool Landscape, double Top, double Bottom, double Left, double Right);

    public PageInfo GetPageSetup()
    {
        var sp = SectionFor(Caret.P); var sz = sp.GetFirstChild<PageSize>(); var m = sp.GetFirstChild<PageMargin>();
        return new PageInfo((sz?.Width?.Value ?? 11906) / 20.0, (sz?.Height?.Value ?? 16838) / 20.0, sz?.Orient?.Value == PageOrientationValues.Landscape,
            (m?.Top?.Value ?? 1440) / 20.0, (m?.Bottom?.Value ?? 1440) / 20.0, (m?.Left?.Value ?? 1440) / 20.0, (m?.Right?.Value ?? 1440) / 20.0);
    }

    /// <summary>Applica a tutte le sezioni (allSections) o solo a quella del caret. Dimensioni in punti, w/h già nell'orientamento voluto.</summary>
    public void SetPageSetup(PageInfo pi, bool allSections = true)
    {
        if (pi.Width < 72 || pi.Height < 72 || pi.Top < 0 || pi.Bottom < 0 || pi.Left < 0 || pi.Right < 0 || pi.Left + pi.Right >= pi.Width - 36 || pi.Top + pi.Bottom >= pi.Height - 36)
            throw new ArgumentException("Dimensioni o margini di pagina non validi.");
        Begin();
        var targets = allSections ? AllSections().ToList() : [SectionFor(Caret.P)];
        foreach (var sp in targets)
        {
            double w = pi.Landscape ? Math.Max(pi.Width, pi.Height) : Math.Min(pi.Width, pi.Height), h = pi.Landscape ? Math.Min(pi.Width, pi.Height) : Math.Max(pi.Width, pi.Height);
            var sz = sp.GetFirstChild<PageSize>() ?? sp.AppendChild(new PageSize());
            sz.Width = (uint)Math.Round(w * 20); sz.Height = (uint)Math.Round(h * 20);
            sz.Orient = pi.Landscape ? PageOrientationValues.Landscape : null;
            var m = sp.GetFirstChild<PageMargin>() ?? sp.AppendChild(new PageMargin { Header = 708, Footer = 708, Gutter = 0 });
            m.Top = (int)Math.Round(pi.Top * 20); m.Bottom = (int)Math.Round(pi.Bottom * 20); m.Left = (uint)Math.Round(pi.Left * 20); m.Right = (uint)Math.Round(pi.Right * 20);
        }
        Changed();
    }

    /// <summary>Numero di colonne (uguali) della sezione del caret.</summary>
    public void SetColumns(int n)
    {
        if (n < 1 || n > 6) throw new ArgumentOutOfRangeException(nameof(n), "Colonne da 1 a 6.");
        Begin();
        var sp = SectionFor(Caret.P);
        var cols = sp.GetFirstChild<Columns>() ?? sp.AppendChild(new Columns());
        cols.ColumnCount = (short)n; cols.Space ??= "720"; cols.EqualWidth = true;
        foreach (var c in cols.Elements<Column>().ToList()) c.Remove();
        // ordine schema: dopo pgMar/paperSrc/pgBorders/lnNumType/pgNumType, prima di docGrid: riposiziona
        cols.Remove();
        var after = sp.Elements().LastOrDefault(e => e is PageMargin or PaperSource or PageBorders or LineNumberType or PageNumberType) ?? sp.Elements().LastOrDefault(e => e is PageSize or HeaderReference or FooterReference or FootnoteProperties or EndnoteProperties or SectionType);
        if (after != null) after.InsertAfterSelf(cols); else sp.PrependChild(cols);
        Changed();
    }

    /// <summary>Interruzione di sezione al caret: pagina successiva o continua.</summary>
    public void InsertSectionBreak(bool continuous)
    {
        Begin();
        if (HasSelection) DeleteSelectionCore();
        var cur = SectionFor(Caret.P);
        var np = XmlOps.Split(Caret.P, Caret.Off);          // il secondo resta dopo
        var first = Caret.P == np ? np : Caret.P;
        var brk = new Paragraph(new ParagraphProperties((SectionProperties)cur.CloneNode(true)));
        var clone = brk.ParagraphProperties!.SectionProperties!;
        clone.GetFirstChild<SectionType>()?.Remove();
        first.InsertAfterSelf(brk);
        // la sezione successiva (quella originale) comincia con il tipo richiesto
        var t = cur.GetFirstChild<SectionType>() ?? cur.PrependChild(new SectionType());
        t.Val = continuous ? SectionMarkValues.Continuous : SectionMarkValues.NextPage;
        // schema: SectionType dopo header/footer reference
        t.Remove();
        var lastRef = cur.Elements().LastOrDefault(e => e is HeaderReference or FooterReference or FootnoteProperties or EndnoteProperties);
        if (lastRef != null) lastRef.InsertAfterSelf(t); else cur.PrependChild(t);
        Set(new Pos(np, 0));
        Changed();
    }

    /// <summary>Testo della storia (intestazione/piè) in modifica.</summary>
    public string GetHeaderFooterTextFromStory() => _storyRoot == null ? "" : string.Join(((char)10).ToString(), _storyRoot.Elements<Paragraph>().Select(ParaText.Text));

    public string GetHeaderFooterText(bool header)
    {
        var sp = SectionFor(Caret.P);
        var part = FindHF(sp, header);
        OpenXmlElement? root = header ? (part as HeaderPart)?.Header : (OpenXmlElement?)(part as FooterPart)?.Footer;
        return root == null ? "" : string.Join("\n", root.Elements<Paragraph>().Select(ParaText.Text));
    }

    OpenXmlPart? FindHF(SectionProperties sp, bool header)
    {
        var main = Doc.Pkg.MainDocumentPart!;
        string? id = header ? sp.Elements<HeaderReference>().FirstOrDefault(h => h.Type?.Value == HeaderFooterValues.Default)?.Id?.Value
                            : sp.Elements<FooterReference>().FirstOrDefault(h => h.Type?.Value == HeaderFooterValues.Default)?.Id?.Value;
        if (id == null) return null;
        try { return main.GetPartById(id); } catch (ArgumentOutOfRangeException) { return null; }
    }

    public void SetHeaderFooter(bool header, string text, Align align, PageNumFormat pn)
    {
        Begin();
        var main = Doc.Pkg.MainDocumentPart!;
        var hfStyle = BuiltIns.EnsureStyle(Doc.Pkg, header ? "Header" : "Footer");
        var sp = SectionFor(Caret.P);
        var part = FindHF(sp, header);
        if (part == null)
        {
            if (header)
            {
                var hp = main.AddNewPart<HeaderPart>(); hp.Header = new Header(); part = hp;
                sp.PrependChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(hp) });
            }
            else
            {
                var fp = main.AddNewPart<FooterPart>(); fp.Footer = new Footer(); part = fp;
                var refEl = new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(fp) };
                var lastHdr = sp.Elements<HeaderReference>().LastOrDefault();
                if (lastHdr != null) lastHdr.InsertAfterSelf(refEl); else sp.PrependChild(refEl);
            }
        }
        OpenXmlCompositeElement root = header ? ((HeaderPart)part).Header! : ((FooterPart)part).Footer!;
        root.RemoveAllChildren();
        var para = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = hfStyle }, new Justification { Val = align == Align.Center ? JustificationValues.Center : align == Align.Right ? JustificationValues.Right : JustificationValues.Left }));
        if (pn == PageNumFormat.PageOfTotal) text = (text.Length > 0 ? text + " " : "") + "Pagina ";
        if (text.Length > 0) para.Append(XmlOps.MakeRun(text, null));
        void Field(string instr, string shown)
        {
            para.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            para.Append(new Run(new FieldCode(" " + instr + " ") { Space = SpaceProcessingModeValues.Preserve }));
            para.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            para.Append(XmlOps.MakeRun(shown, null));
            para.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }
        if (pn != PageNumFormat.None) { Field("PAGE", "1"); if (pn == PageNumFormat.PageOfTotal) { para.Append(XmlOps.MakeRun(" di ", null)); Field("NUMPAGES", "1"); } }
        root.Append(para);
        Changed();
    }

    // ---------- trova e sostituisci ----------
    public Pos? Find(string what, bool matchCase, bool fromCaret = true)
    {
        if (what.Length == 0) return null;
        var ps = Paragraphs(); var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        int i0 = fromCaret ? Idx(Range.End.P) : 0, o0 = fromCaret ? Range.End.Off : 0;
        for (int n = 0; n <= ps.Count; n++)
        {
            int i = (i0 + n) % ps.Count;
            var t = ParaText.Text(ps[i]);
            int from = n == 0 ? Math.Min(o0, t.Length) : 0;
            if (n == ps.Count) { t = t[..Math.Min(t.Length, o0 + what.Length)]; from = 0; }
            int k = t.IndexOf(what, from, cmp);
            if (k >= 0) return new Pos(ps[i], k);
        }
        return null;
    }

    public bool FindNext(string what, bool matchCase)
    {
        var f = Find(what, matchCase);
        if (f == null) return false;
        Anchor = f.Value; Caret = new Pos(f.Value.P, f.Value.Off + what.Length); Pending = null;
        return true;
    }

    public int ReplaceAll(string what, string with, bool matchCase)
    {
        if (what.Length == 0) return 0;
        Begin();
        int count = 0;
        var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        foreach (var p in Paragraphs().ToList())
        {
            var t = ParaText.Text(p);
            var hits = new List<int>(); int k = 0;
            while ((k = t.IndexOf(what, k, cmp)) >= 0) { hits.Add(k); k += what.Length; }
            for (int h = hits.Count - 1; h >= 0; h--)
            {
                var leaves = ParaText.Leaves(p);
                var lf = leaves.FirstOrDefault(l => l.Kind == LeafKind.Text && l.Start <= hits[h] && hits[h] < l.End);
                var rpr = (RunProperties?)lf?.Run.RunProperties?.CloneNode(true);
                XmlOps.DeleteRange(p, hits[h], hits[h] + what.Length);
                XmlOps.InsertText(p, hits[h], with, rpr);
                count++;
            }
        }
        var ps = Paragraphs();
        Set(new Pos(ps.Contains(Caret.P) ? Caret.P : ps[0], Math.Min(Caret.Off, ParaText.Length(ps.Contains(Caret.P) ? Caret.P : ps[0]))));
        if (count > 0) Changed();
        return count;
    }
}
