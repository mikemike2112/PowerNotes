using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace PowerNotes.Core.Model;

public enum Align { Left, Center, Right, Justify }

public sealed record TabDef(double Pos, string Kind, string Leader);

/// <summary>Formattazione effettiva di un run (risultato della cascata di stili).</summary>
public sealed class RunFmt
{
    public string Font = "Calibri";
    public double Size = 10;
    public bool Bold, Italic, Strike, SmallCaps, AllCaps;
    public bool Underline;
    public string? Color;        // "RRGGBB" oppure null = automatico
    public string? Highlight;    // nome OOXML (yellow, green…) oppure null
    public int VertAlign;        // 0 normale, 1 apice, -1 pedice
    public RunFmt Clone() => (RunFmt)MemberwiseClone();
    /// <summary>Corpo effettivo: apici/pedici e maiuscoletto riducono la dimensione.</summary>
    public double EffSize => VertAlign != 0 ? Size * 0.65 : Size;
}

/// <summary>Formattazione effettiva di un paragrafo.</summary>
public sealed class ParaFmt
{
    public Align Alignment;
    public double IndLeft, IndRight, FirstLine;       // FirstLine negativo = sporgente
    public double Before, After;
    public double Line = 240;                          // twips o 240esimi
    public string LineRule = "auto";
    public bool KeepNext, KeepLines, PageBreakBefore, ContextualSpacing;
    public bool WidowControl = true;
    public List<TabDef> Tabs = new();
    public int? NumId, Ilvl;
    public string? StyleId;
    public int OutlineLevel = 9;
    public ParaFmt Clone() { var c = (ParaFmt)MemberwiseClone(); c.Tabs = new(Tabs); return c; }
}

public sealed class NumLevel
{
    public string Fmt = "decimal", Text = "%1.", Suff = "tab", Font = "";
    public int Start = 1;
    public Action<RunFmt>? ApplyRun;
}

/// <summary>Risolve gli stili: docDefaults → (stile tabella) → stile paragrafo → numerazione → stile carattere → formattazione diretta.</summary>
public sealed class StyleResolver
{
    readonly Styles? _styles;
    readonly Numbering? _numbering;
    readonly Dictionary<string, Style> _byId = new();
    readonly Dictionary<string, (ParaFmt, RunFmt)> _paraCache = new();
    readonly string _minor = "Calibri", _major = "Calibri Light";
    public double DefaultTab = 36;
    Style? _defaultPara, _defaultTable;

    public StyleResolver(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        _styles = main?.StyleDefinitionsPart?.Styles;
        _numbering = main?.NumberingDefinitionsPart?.Numbering;
        var fs = main?.ThemePart?.Theme?.ThemeElements?.FontScheme;
        _minor = fs?.MinorFont?.LatinFont?.Typeface?.Value is { Length: > 0 } mn ? mn : _minor;
        _major = fs?.MajorFont?.LatinFont?.Typeface?.Value is { Length: > 0 } mj ? mj : _major;
        var dt = main?.DocumentSettingsPart?.Settings?.GetFirstChild<DefaultTabStop>()?.Val?.Value;
        if (dt != null) DefaultTab = dt.Value / 20.0;
        if (_styles != null)
            foreach (var s in _styles.Elements<Style>())
            {
                if (s.StyleId?.Value is { } id) _byId[id] = s;
                if (s.Default?.Value == true)
                {
                    if (s.Type?.Value == StyleValues.Paragraph) _defaultPara = s;
                    if (s.Type?.Value == StyleValues.Table) _defaultTable = s;
                }
            }
    }

    public Style? Find(string? id) => id != null && _byId.TryGetValue(id, out var s) ? s : null;

    public string? DefaultTableStyleId => _defaultTable?.StyleId?.Value;

    public IEnumerable<Style> Chain(string? id)
    {
        var list = new List<Style>();
        var seen = new HashSet<string>();
        for (var s = Find(id); s != null && seen.Add(s.StyleId?.Value ?? ""); s = Find(s.BasedOn?.Val?.Value)) list.Insert(0, s);
        return list;
    }

    // ---- run ----
    public void ApplyRPr(RunFmt r, OpenXmlElement? pr)
    {
        if (pr == null) return;
        foreach (var e in pr.Elements())
            switch (e)
            {
                case RunFonts f:
                    var th = f.AsciiTheme?.Value;
                    if (th != null) r.Font = th.ToString()!.StartsWith("Major", StringComparison.OrdinalIgnoreCase) ? _major : _minor;
                    else if (f.Ascii?.Value is { Length: > 0 } a) r.Font = a;
                    else if (f.HighAnsi?.Value is { Length: > 0 } h) r.Font = h;
                    break;
                case Bold b: r.Bold = b.Val?.Value ?? true; break;
                case Italic i: r.Italic = i.Val?.Value ?? true; break;
                case Strike s: r.Strike = s.Val?.Value ?? true; break;
                case DoubleStrike ds: r.Strike = ds.Val?.Value ?? true; break;
                case SmallCaps sc: r.SmallCaps = sc.Val?.Value ?? true; break;
                case Caps c: r.AllCaps = c.Val?.Value ?? true; break;
                case Underline u: r.Underline = u.Val?.Value != UnderlineValues.None; break;
                case Color c: r.Color = c.Val?.Value is { } v && v != "auto" ? v : null; break;
                case FontSize fs: if (double.TryParse(fs.Val?.Value, System.Globalization.CultureInfo.InvariantCulture, out var hp)) r.Size = hp / 2; break;
                case Highlight hl: r.Highlight = hl.Val?.Value is { } hv && hv != HighlightColorValues.None ? hl.Val.InnerText : null; break;
                case VerticalTextAlignment va:
                    r.VertAlign = va.Val?.Value == VerticalPositionValues.Superscript ? 1 : va.Val?.Value == VerticalPositionValues.Subscript ? -1 : 0;
                    break;
            }
    }

    // ---- paragrafo ----
    public static void ApplyPPr(ParaFmt p, OpenXmlElement? pr)
    {
        if (pr == null) return;
        foreach (var e in pr.Elements())
            switch (e)
            {
                case Justification j:
                    var jv = j.Val?.Value;
                    p.Alignment = jv == JustificationValues.Center ? Align.Center : jv == JustificationValues.Right || jv == JustificationValues.End ? Align.Right
                        : jv == JustificationValues.Both || jv == JustificationValues.Distribute ? Align.Justify : Align.Left;
                    break;
                case Indentation ind:
                    if (Tw(ind.Left?.Value ?? ind.Start?.Value) is { } l) p.IndLeft = l;
                    if (Tw(ind.Right?.Value ?? ind.End?.Value) is { } r) p.IndRight = r;
                    if (Tw(ind.FirstLine?.Value) is { } fl) p.FirstLine = fl;
                    if (Tw(ind.Hanging?.Value) is { } hg) p.FirstLine = -hg;
                    break;
                case SpacingBetweenLines sp:
                    if (Tw(sp.Before?.Value) is { } bf) p.Before = bf;
                    if (Tw(sp.After?.Value) is { } af) p.After = af;
                    if (double.TryParse(sp.Line?.Value, out var ln)) p.Line = ln;
                    if (sp.LineRule?.Value is { } lr) p.LineRule = lr == LineSpacingRuleValues.Exact ? "exact" : lr == LineSpacingRuleValues.AtLeast ? "atLeast" : "auto";
                    break;
                case KeepNext k: p.KeepNext = k.Val?.Value ?? true; break;
                case KeepLines k: p.KeepLines = k.Val?.Value ?? true; break;
                case PageBreakBefore k: p.PageBreakBefore = k.Val?.Value ?? true; break;
                case WidowControl k: p.WidowControl = k.Val?.Value ?? true; break;
                case ContextualSpacing k: p.ContextualSpacing = k.Val?.Value ?? true; break;
                case OutlineLevel o: p.OutlineLevel = o.Val?.Value ?? 9; break;
                case NumberingProperties np:
                    if (np.NumberingId?.Val?.Value is { } nid) p.NumId = nid;
                    if (np.NumberingLevelReference?.Val?.Value is { } il) p.Ilvl = il;
                    break;
                case Tabs tabs:
                    foreach (var t in tabs.Elements<DocumentFormat.OpenXml.Wordprocessing.TabStop>())
                    {
                        var kind = t.Val?.InnerText ?? "left";
                        double pos = (t.Position?.Value ?? 0) / 20.0;
                        p.Tabs.RemoveAll(x => Math.Abs(x.Pos - pos) < 0.01);
                        if (kind != "clear") p.Tabs.Add(new TabDef(pos, kind, t.Leader?.InnerText ?? "none"));
                    }
                    break;
            }
    }

    static double? Tw(string? v) => double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out var d) ? d / 20.0 : null;
    static double? Tw(int? v) => v == null ? null : v.Value / 20.0;

    (ParaFmt, RunFmt) Defaults()
    {
        var p = new ParaFmt(); var r = new RunFmt();
        var dd = _styles?.DocDefaults;
        ApplyRPr(r, dd?.RunPropertiesDefault?.RunPropertiesBaseStyle);
        ApplyPPr(p, dd?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle);
        return (p, r);
    }

    /// <summary>Formattazione di base di un paragrafo con dato stile (e stile tabella se dentro tabella).</summary>
    public (ParaFmt, RunFmt) ParaBase(string? styleId, string? tableStyleId = null)
    {
        var key = (styleId ?? "") + "|" + (tableStyleId ?? "");
        if (_paraCache.TryGetValue(key, out var c)) return (c.Item1.Clone(), c.Item2.Clone());
        var (p, r) = Defaults();
        if (tableStyleId != null || _defaultTable != null)
            foreach (var s in Chain(tableStyleId ?? _defaultTable?.StyleId?.Value)) { ApplyPPr(p, s.StyleParagraphProperties); ApplyRPr(r, s.StyleRunProperties); }
        var sid = styleId ?? _defaultPara?.StyleId?.Value;
        p.StyleId = sid;
        foreach (var s in Chain(sid))
        {
            ApplyPPr(p, s.StyleParagraphProperties);
            ApplyRPr(r, s.StyleRunProperties);
            // numerazione legata allo stile senza ilvl: cerca il livello con pStyle
            if (s.StyleParagraphProperties?.NumberingProperties is { } np && np.NumberingLevelReference == null && np.NumberingId?.Val?.Value is { } nid)
                p.Ilvl = FindLevelForStyle(nid, s.StyleId?.Value) ?? 0;
        }
        _paraCache[key] = (p.Clone(), r.Clone());
        return (p, r);
    }

    public ParaFmt ResolvePara(Paragraph para, string? tableStyleId, out RunFmt markRun) => ResolvePara(para, tableStyleId, out markRun, out _);

    /// <summary>markRun = segno di paragrafo (stile + pPr/rPr); baseRun = base per i run (solo stile).</summary>
    public ParaFmt ResolvePara(Paragraph para, string? tableStyleId, out RunFmt markRun, out RunFmt baseRun)
    {
        var pPr = para.ParagraphProperties;
        var (p, r) = ParaBase(pPr?.ParagraphStyleId?.Val?.Value, tableStyleId);
        // numerazione: id dallo stile o diretto, poi pPr del livello, poi diretto
        var direct = pPr?.NumberingProperties;
        int? numId = direct?.NumberingId?.Val?.Value ?? p.NumId;
        int ilvl = direct?.NumberingLevelReference?.Val?.Value ?? p.Ilvl ?? 0;
        p.NumId = numId; p.Ilvl = ilvl;
        if (numId is > 0 && GetLevelPPr(numId.Value, ilvl) is { } lp) ApplyPPr(p, lp);
        ApplyPPr(p, pPr);
        p.NumId = numId; p.Ilvl = ilvl;
        baseRun = r.Clone();
        markRun = r.Clone();
        ApplyRPr(markRun, pPr?.ParagraphMarkRunProperties);
        return p;
    }

    public RunFmt ResolveRun(RunFmt paraRun, Run run)
    {
        var r = paraRun.Clone();
        var rPr = run.RunProperties;
        foreach (var s in Chain(rPr?.RunStyle?.Val?.Value)) ApplyRPr(r, s.StyleRunProperties);
        ApplyRPr(r, rPr);
        return r;
    }

    // ---- numerazione ----
    int? FindLevelForStyle(int numId, string? styleId)
    {
        var abs = Abstract(numId);
        if (abs == null || styleId == null) return null;
        foreach (var l in abs.Elements<Level>()) if (l.ParagraphStyleIdInLevel?.Val?.Value == styleId) return l.LevelIndex?.Value;
        return null;
    }

    AbstractNum? Abstract(int numId)
    {
        var num = _numbering?.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numId);
        var aid = num?.AbstractNumId?.Val?.Value;
        return aid == null ? null : _numbering!.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == aid);
    }

    Level? RawLevel(int numId, int ilvl)
    {
        var num = _numbering?.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numId);
        var ov = num?.Elements<LevelOverride>().FirstOrDefault(o => o.LevelIndex?.Value == ilvl)?.Level;
        return ov ?? Abstract(numId)?.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == ilvl);
    }

    OpenXmlElement? GetLevelPPr(int numId, int ilvl) => RawLevel(numId, ilvl)?.PreviousParagraphProperties;

    public NumLevel? GetLevel(int numId, int ilvl)
    {
        var l = RawLevel(numId, ilvl);
        if (l == null) return null;
        var start = l.StartNumberingValue?.Val?.Value ?? 1;
        var so = _numbering?.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numId)?.Elements<LevelOverride>().FirstOrDefault(o => o.LevelIndex?.Value == ilvl)?.StartOverrideNumberingValue?.Val?.Value;
        var nl = new NumLevel
        {
            Fmt = l.NumberingFormat?.Val?.InnerText ?? "decimal",
            Text = l.LevelText?.Val?.Value ?? "",
            Suff = l.LevelSuffix?.Val?.InnerText ?? "tab",
            Start = so ?? start
        };
        var pr = l.NumberingSymbolRunProperties;
        nl.ApplyRun = rf => ApplyRPr(rf, pr);
        nl.Font = pr?.GetFirstChild<RunFonts>()?.Ascii?.Value ?? "";
        return nl;
    }
}
