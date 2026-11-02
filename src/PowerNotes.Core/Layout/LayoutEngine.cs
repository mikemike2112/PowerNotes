using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Fonts;
using PowerNotes.Core.Model;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace PowerNotes.Core.Layout;

/// <summary>Motore di layout: stile effettivo → line breaking con metriche reali → paginazione. Una sola pipeline per schermo, PDF e stampa.</summary>
public sealed class LayoutEngine
{
    // ---- strutture intermedie ----
    abstract class Block { public double Height; }
    sealed class Atom { public char K; /* w parola, s spazi, t tab, b a capo, p page break, i immagine */ public string Text = ""; public RunFmt Fmt = new(); public double W; public int Start, Len; public bool Synth; public string? MetricFont; public ImagePart? Img; public double ImgH; public int Spaces; public char TabKind = 'l'; public long Note; /* id nota a piè di pagina di cui questo atomo è il richiamo */ }
    sealed class SegL { public Atom A = null!; public double X; public double W; }
    sealed class LineL { public double NatH; public double H, Base; public List<SegL> Segs = new(); public int Start, End; public bool PageBreakAfter, Last; public double Left, Right; }
    sealed class ParaL : Block { public List<FloatL> Floats = new(); public Story Story; public bool Collapsed; public Paragraph P = null!; public ParaFmt F = null!; public List<LineL> Lines = new(); public double Before, After, X, Width; }
    sealed class CellL { public double X, W, ContentH, H; public int RowSpan = 1, Col, Span; public List<Block> Content = new(); public string? Fill; public Bdr?[] B = new Bdr?[4]; public bool Continue; public string VAlign = "top"; public double MarT, MarB; }
    sealed class RowL { public bool IsHeader; public TableRow? Src; public double H; public List<CellL> Cells = new(); }
    sealed class TableL : Block { public double X; public List<RowL> Rows = new(); public double[] Grid = []; }
    record Bdr(double W, string Color);
    /// <summary>Oggetto ancorato (immagine o casella di testo): posizione risolta in fase di emissione.</summary>
    sealed class FloatL { public double W, H, OffX, OffY, Dist; public string HFrom = "column", VFrom = "paragraph"; public string? HAlign, VAlign, Fill, LineColor; public double LineW; public bool Behind; public string Wrap = "none"; public ImagePart? Img; public List<Paragraph>? BoxParas; }

    sealed class Ctx
    {
        public required StyleResolver St;
        public required OpenXmlPart Part;
        public int PageNo = 1, NumPages = 1;
        public Dictionary<int, int[]> Counters = new();
        public Dictionary<Run, RunFmt> RunCache = new();
        public Dictionary<long, int>? Notes;   // solo corpo: id nota → numero progressivo
        public int MarkNo;                      // dentro una nota: numero del suo richiamo
        public Action<RunFmt>? CondRun;   // formattazione condizionale dello stile tabella (prima riga in grassetto, ecc.)
    }

    sealed class SectInfo
    {
        public SectionProperties? Pr;
        public double W = 612, H = 792, Top = 72, Bottom = 72, Left = 72, Right = 72, Header = 36, Footer = 36, Gutter;
        public bool TitlePg, Continuous;
        public int Cols = 1; public double ColGap = 36;
        public double ColW => Cols > 1 ? (ContentW - ColGap * (Cols - 1)) / Cols : ContentW;
        public OpenXmlPart? HdrDefault, HdrFirst, FtrDefault, FtrFirst;
        public double ContentW => W - Left - Right - Gutter;
    }

    readonly WordprocessingDocument _doc;
    readonly StyleResolver _st;
    readonly LayoutResult _res = new();
    readonly bool _compat15;
    Ctx _body = null!;
    // stato di paginazione
    Page _pg = null!;
    double _lastLineY, _colShift;
    int _col;
    double _y, _bottom, _top, _left, _width;
    bool _pageHasBody, _hardTop;
    SectInfo _sec = null!;
    int _secPageIdx;
    readonly List<(Page pg, SectInfo sec, int idxInSec)> _pageSec = new();
    // note a piè di pagina. ponytail: solo note nei paragrafi del corpo (non in tabelle), niente note di chiusura → estendere CommitNotes ai blocchi di cella
    const double NoteSep = 8;
    Dictionary<long, Footnote>? _noteSrc;
    readonly Dictionary<long, List<Block>> _noteCache = new();
    readonly List<long> _pageNotes = new();
    double _noteTotal, _pageBottom;

    LayoutEngine(WordprocessingDocument doc)
    {
        _doc = doc;
        _st = new StyleResolver(doc);
        var cs = doc.MainDocumentPart?.DocumentSettingsPart?.Settings?.GetFirstChild<Compatibility>()?.Elements<CompatibilitySetting>().FirstOrDefault(c => c.Name?.InnerText == "compatibilityMode");
        _compat15 = cs?.Val?.Value is { } v && int.TryParse(v, out var n) ? n >= 15 : false;
    }

    public static LayoutResult Layout(WordprocessingDocument doc)
    {
        var e = new LayoutEngine(doc);
        e.Run();
        return e._res;
    }

    // ================= pagine / sezioni =================
    void Run()
    {
        var main = _doc.MainDocumentPart ?? throw new InvalidDataException("Documento senza parte principale.");
        var body = main.Document?.Body;
        _body = new Ctx { St = _st, Part = main, Notes = new() };
        var sections = new List<(SectInfo sec, List<OpenXmlElement> blocks)>();
        var cur = new List<OpenXmlElement>();
        SectInfo? prev = null;
        void Close(SectionProperties? pr)
        {
            var si = ReadSect(pr, prev);
            sections.Add((si, cur)); cur = new(); prev = si;
        }
        if (body != null)
            foreach (var el in Flatten(body.ChildElements))
            {
                if (el is Paragraph p) { cur.Add(p); if (p.ParagraphProperties?.SectionProperties is { } sp) Close(sp); }
                else if (el is Table) cur.Add(el);
            }
        Close(body?.Elements<SectionProperties>().LastOrDefault());

        bool first = true;
        foreach (var (si, blocks) in sections)
        {
            _sec = si;
            bool contiguous = !first && si.Continuous;
            if (!contiguous) NewPage(first ? false : true, hard: true, first: true);
            else { _left = si.Left + si.Gutter; _width = si.ColW; _col = 0; _colShift = 0; }
            first = false;
            var built = new List<Block>();
            foreach (var b in blocks)
                built.Add(b is Paragraph p ? BuildPara(p, _left, _width, null, _body) : BuildTable((Table)b, _left, _width, _body));
            PlaceBlocks(built);
        }
        FlushNotes();
        if (_res.Pages.Count == 0) { _sec = ReadSect(null, null); NewPage(false, true, true); }
        // intestazioni e piè di pagina: ora si conosce il totale
        int total = _res.Pages.Count;
        foreach (var (pg, sec, idx) in _pageSec) DrawHeaderFooter(pg, sec, idx, total);
    }

    static IEnumerable<OpenXmlElement> Flatten(IEnumerable<OpenXmlElement> els)
    {
        foreach (var e in els)
        {
            if (e is SdtBlock sb) { foreach (var x in Flatten(sb.SdtContentBlock?.ChildElements ?? Enumerable.Empty<OpenXmlElement>())) yield return x; }
            else if (e is CustomXmlBlock cb) { foreach (var x in Flatten(cb.ChildElements)) yield return x; }
            else yield return e;
        }
    }

    SectInfo ReadSect(SectionProperties? pr, SectInfo? prev)
    {
        var s = new SectInfo { Pr = pr };
        if (prev != null) { s.HdrDefault = prev.HdrDefault; s.HdrFirst = prev.HdrFirst; s.FtrDefault = prev.FtrDefault; s.FtrFirst = prev.FtrFirst; }
        if (pr == null) return s;
        var sz = pr.GetFirstChild<PageSize>();
        if (sz?.Width?.Value is { } w) s.W = w / 20.0;
        if (sz?.Height?.Value is { } h) s.H = h / 20.0;
        if (pr.GetFirstChild<PageMargin>() is { } m)
        {
            s.Top = (m.Top?.Value ?? 1440) / 20.0; s.Bottom = (m.Bottom?.Value ?? 1440) / 20.0;
            s.Left = (m.Left?.Value ?? 1440) / 20.0; s.Right = (m.Right?.Value ?? 1440) / 20.0;
            s.Header = (m.Header?.Value ?? 720) / 20.0; s.Footer = (m.Footer?.Value ?? 720) / 20.0; s.Gutter = (m.Gutter?.Value ?? 0) / 20.0;
        }
        s.TitlePg = pr.GetFirstChild<TitlePage>() is { } tp && (tp.Val?.Value ?? true);
        if (pr.GetFirstChild<Columns>() is { } cols) { s.Cols = Math.Clamp((int)(cols.ColumnCount?.Value ?? 1), 1, 12); s.ColGap = (double.TryParse(cols.Space?.Value, out var csp) ? csp : 720) / 20.0; }
        s.Continuous = pr.GetFirstChild<SectionType>()?.Val?.Value == SectionMarkValues.Continuous;
        var main = _doc.MainDocumentPart!;
        foreach (var hr in pr.Elements<HeaderReference>())
        {
            var part = TryPart(main, hr.Id?.Value);
            if (hr.Type?.Value == HeaderFooterValues.First) s.HdrFirst = part; else if (hr.Type?.Value == HeaderFooterValues.Default) s.HdrDefault = part;
        }
        foreach (var fr in pr.Elements<FooterReference>())
        {
            var part = TryPart(main, fr.Id?.Value);
            if (fr.Type?.Value == HeaderFooterValues.First) s.FtrFirst = part; else if (fr.Type?.Value == HeaderFooterValues.Default) s.FtrDefault = part;
        }
        return s;
    }

    static OpenXmlPart? TryPart(OpenXmlPart owner, string? id)
    {
        if (id == null) return null;
        try { return owner.GetPartById(id); } catch (ArgumentOutOfRangeException) { return null; }
    }

    OpenXmlPart? HdrFor(SectInfo s, int idx) => s.TitlePg && idx == 0 ? s.HdrFirst : s.HdrDefault;
    OpenXmlPart? FtrFor(SectInfo s, int idx) => s.TitlePg && idx == 0 ? s.FtrFirst : s.FtrDefault;

    List<Block> BuildHF(OpenXmlPart? part, SectInfo s, int pageNo, int numPages)
    {
        var list = new List<Block>();
        OpenXmlElement? root = (part as HeaderPart)?.Header ?? (OpenXmlElement?)(part as FooterPart)?.Footer;
        if (root == null) return list;
        var c = new Ctx { St = _st, Part = part!, PageNo = pageNo, NumPages = numPages };
        foreach (var e in Flatten(root.ChildElements))
            if (e is Paragraph p) list.Add(BuildPara(p, s.Left + s.Gutter, s.ContentW, null, c));
            else if (e is Table t) list.Add(BuildTable(t, s.Left + s.Gutter, s.ContentW, c));
        var story = part is HeaderPart ? Story.Header : Story.Footer;
        void Mark(IEnumerable<Block> bs) { foreach (var b in bs) { if (b is ParaL pl) pl.Story = story; else if (b is TableL tl) foreach (var r in tl.Rows) foreach (var ce in r.Cells) Mark(ce.Content); } }
        Mark(list);
        return list;
    }

    double StackHeight(List<Block> bs)
    {
        double h = 0; ParaL? prev = null;
        foreach (var b in bs)
        {
            if (b is ParaL p) { h += Gap(prev, p) + p.Lines.Sum(l => l.H) + 0; prev = p; h += 0; }
            else { h += b.Height; prev = null; }
            if (b is ParaL pl) h += pl.After;
        }
        return h;
    }

    /// <summary>Spazio verticale da aggiungere prima di p, al netto dell'After di prev (già aggiunto). Word usa il MAGGIORE fra After e Before; spaziatura contestuale: stesso stile → zero.</summary>
    static double Gap(ParaL? prev, ParaL p)
    {
        if (prev == null) return p.Before;
        bool same = prev.F.StyleId == p.F.StyleId;
        double pa = same && prev.F.ContextualSpacing ? 0 : prev.After, pb = same && p.F.ContextualSpacing ? 0 : p.Before;
        return Math.Max(pa, pb) - prev.After;
    }

    void NewPage(bool addPage, bool hard, bool first)
    {
        if (!first && !hard && _sec.Cols > 1 && _col < _sec.Cols - 1)
        {
            _col++; _colShift = _col * (_sec.ColW + _sec.ColGap); _y = _top; _pageHasBody = false; _hardTop = false;   // colonna successiva sulla stessa pagina
            return;
        }
        FlushNotes();
        _col = 0; _colShift = 0;
        _secPageIdx = first ? 0 : _secPageIdx + 1;
        _pg = new Page { Index = _res.Pages.Count, Number = _res.Pages.Count + 1, Width = _sec.W, Height = _sec.H, MarL = _sec.Left + _sec.Gutter, MarR = _sec.Right, MarT = _sec.Top, MarB = _sec.Bottom };
        _res.Pages.Add(_pg);
        _pageSec.Add((_pg, _sec, _secPageIdx));
        double hdrH = StackHeight(BuildHF(HdrFor(_sec, _secPageIdx), _sec, 1, 1));
        double ftrH = StackHeight(BuildHF(FtrFor(_sec, _secPageIdx), _sec, 1, 1));
        _top = Math.Max(_sec.Top, hdrH > 0 ? _sec.Header + hdrH : 0);
        _bottom = Math.Min(_sec.H - _sec.Bottom, ftrH > 0 ? _sec.H - _sec.Footer - ftrH : double.MaxValue);
        _pageBottom = _bottom;
        _left = _sec.Left + _sec.Gutter; _width = _sec.ColW;
        _y = _top; _pageHasBody = false; _hardTop = hard;
    }

    void DrawHeaderFooter(Page pg, SectInfo s, int idx, int total)
    {
        var save = _pg; _pg = pg; var saveShift = _colShift; _colShift = 0;
        var hb = BuildHF(HdrFor(s, idx), s, pg.Number, total);
        EmitStack(hb, s.Header);
        var fb = BuildHF(FtrFor(s, idx), s, pg.Number, total);
        EmitStack(fb, s.H - s.Footer - StackHeight(fb));
        _pg = save; _colShift = saveShift;
    }

    void EmitStack(List<Block> bs, double y)
    {
        ParaL? prev = null;
        foreach (var b in bs)
        {
            if (b is ParaL p) { y += Gap(prev, p); EmitPara(p, 0, p.Lines.Count, y); y += p.Lines.Sum(l => l.H) + p.After; prev = p; }
            else if (b is TableL t) { EmitTable(t, 0, t.Rows.Count, y); y += t.Height; prev = null; }
        }
    }

    // ================= paginazione corpo =================
    void PlaceBlocks(List<Block> blocks)
    {
        ParaL? prev = null;
        for (int bi = 0; bi < blocks.Count; bi++)
        {
            var b = blocks[bi];
            if (b is ParaL p)
            {
                if (p.F.PageBreakBefore && _pageHasBody) NewPage(true, true, false);
                // keep-with-next: il gruppo deve stare insieme alla prima riga del blocco successivo
                if (p.F.KeepNext && _pageHasBody)
                {
                    double need = 0; int j = bi; ParaL? pv = prev;
                    while (j < blocks.Count && blocks[j] is ParaL q && q.F.KeepNext) { need += Gap(pv, q) + q.Lines.Sum(l => l.H) + q.After; pv = q; j++; }
                    if (j < blocks.Count) need += blocks[j] is ParaL n ? Gap(pv, n) + n.Lines.Take(2).Sum(l => l.H) : ((TableL)blocks[j]).Rows.FirstOrDefault()?.H ?? 0;
                    if (_y + need > _bottom + 0.01 && need < _bottom - _top) NewPage(true, false, false);
                }
                PlacePara(p, prev);
                prev = p;
            }
            else
            {
                PlaceTable((TableL)b);
                prev = null;
            }
        }
    }

    void PlacePara(ParaL p, ParaL? prev)
    {
        double before = Gap(prev, p);
        if (!_pageHasBody && !_hardTop && _res.Pages.Count > 1 && _pg.Index > 0) before = 0;   // spazio prima soppresso in cima a pagina dopo salto automatico
        if (prev != null && prev.F.ContextualSpacing && prev.F.StyleId == p.F.StyleId) { /* dopo già azzerato sotto */ }
        _y += before;
        int n = p.Lines.Count, i = 0;
        while (i < n)
        {
            double avail = _bottom - _y;
            int k = 0; double h = 0;
            while (i + k < n && h + p.Lines[i + k].H + Reserve(p, i, i + k + 1) <= avail + 0.01) { h += p.Lines[i + k].H; k++; }
            if (i + k < n)
            {
                int rem = n - i;
                if (p.F.KeepLines && i == 0 && _pageHasBody) k = 0;
                else if (p.F.WidowControl)
                {
                    if (i == 0 && k < 2) k = 0;
                    else if (n - (i + k) < 2) { k = rem - 2; if (i == 0 && k < 2) k = 0; if (k < 0) k = 0; }
                }
                if (k == 0)
                {
                    if (!_pageHasBody) k = 1;    // pagina vuota: evita cicli infiniti
                    else { NewPage(true, false, false); _y = _top; continue; }
                }
            }
            // emetti le righe i..i+k fermandosi a un salto pagina esplicito
            int e = i;
            double yy = _y;
            while (e < i + k)
            {
                yy += p.Lines[e].H; e++;
                if (p.Lines[e - 1].PageBreakAfter) break;
            }
            CommitNotes(p, i, e);
            EmitPara(p, i, e, _y);
            _pageHasBody = true;
            _y = yy;
            bool brk = p.Lines[e - 1].PageBreakAfter;
            i = e;
            if (brk) { NewPage(true, true, false); if (i >= n) { _y = _top; return; } }
            else if (i < n) { NewPage(true, false, false); }
        }
        _y += p.After;
    }

    void PlaceTable(TableL t)
    {
        int n = t.Rows.Count, r = 0;
        int hdr = 0; while (hdr < n && t.Rows[hdr].IsHeader) hdr++;      // righe di intestazione ripetute
        double hdrH = t.Rows.Take(hdr).Sum(x => x.H);
        while (r < n)
        {
            int end = r;
            // gruppo di righe collegate da unioni verticali
            for (int j = r; j <= end; j++) foreach (var c in t.Rows[j].Cells) end = Math.Max(end, j + c.RowSpan - 1);
            double h = 0;
            for (int j = r; j <= end; j++) h += t.Rows[j].H;
            // ponytail: si spezza solo la riga più alta di una pagina intera (le altre vanno alla pagina dopo, come prima) → spezzare ogni riga che non sta, rispettando cantSplit
            if (end == r && h > _bottom - _top + 0.01 && t.Rows[r].Cells.All(x => x.RowSpan == 1 && !x.Continue)) { SplitRow(t, r, hdr, hdrH); r++; continue; }
            if (_y + h > _bottom + 0.01 && _pageHasBody)
            {
                NewPage(true, false, false);
                if (hdr > 0 && r >= hdr && hdrH + h < _bottom - _top) { EmitTable(t, 0, hdr, _y); _y += hdrH; }
            }
            EmitTable(t, r, end + 1, _y);
            _y += h; _pageHasBody = true;
            r = end + 1;
        }
    }

    // ================= note a piè di pagina =================
    Footnote? FindNote(long id)
    {
        _noteSrc ??= (_doc.MainDocumentPart?.FootnotesPart?.Footnotes?.Elements<Footnote>() ?? Enumerable.Empty<Footnote>())
            .Where(f => f.Id?.Value != null && (f.Type?.Value is null || f.Type.Value == FootnoteEndnoteValues.Normal)).GroupBy(f => f.Id!.Value).ToDictionary(g => g.Key, g => g.First());
        return _noteSrc.GetValueOrDefault(id);
    }

    List<Block> NoteBlocks(long id)
    {
        if (_noteCache.TryGetValue(id, out var b)) return b;
        var c = new Ctx { St = _st, Part = _doc.MainDocumentPart!.FootnotesPart!, MarkNo = _body.Notes!.GetValueOrDefault(id) };
        b = new List<Block>();
        foreach (var p in FindNote(id)!.Elements<Paragraph>()) b.Add(BuildPara(p, _sec.Left + _sec.Gutter, _sec.ContentW, null, c));
        return _noteCache[id] = b;
    }

    /// <summary>Atomo sintetico (numero in apice) per il richiamo di nota nel corpo o per il segno di nota dentro la nota.</summary>
    void AddNoteMark(List<Atom> atoms, Run r, long id, RunFmt baseRun, Ctx c)
    {
        int no;
        if (id == -1) no = c.MarkNo;
        else if (c.Notes == null || FindNote(id) == null) return;
        else if (!c.Notes.TryGetValue(id, out no)) c.Notes[id] = no = c.Notes.Count + 1;
        if (no == 0) return;
        var f = c.St.ResolveRun(baseRun, r).Clone(); if (f.VertAlign == 0) f.VertAlign = 1;
        var txt = no.ToString();
        atoms.Add(new Atom { K = 'w', Text = txt, Fmt = f, W = Metrics.Width(txt, f), Synth = true, Note = id == -1 ? 0 : id });
    }

    /// <summary>Altezza da riservare in fondo alla pagina per le note dei richiami nelle righe a..b-1 (più il separatore se è la prima nota della pagina).</summary>
    double Reserve(ParaL p, int a, int b)
    {
        double h = 0; bool any = false;
        for (int i = a; i < b; i++)
            foreach (var s in p.Lines[i].Segs)
                if (s.A.Note != 0) { any = true; h += StackH(NoteBlocks(s.A.Note)); }
        return any && _pageNotes.Count == 0 ? h + NoteSep : h;
    }

    void CommitNotes(ParaL p, int a, int b)
    {
        double h = Reserve(p, a, b);
        if (h == 0) return;
        for (int i = a; i < b; i++) foreach (var s in p.Lines[i].Segs) if (s.A.Note != 0) _pageNotes.Add(s.A.Note);
        _bottom -= h; _noteTotal += h;
    }

    void FlushNotes()
    {
        if (_pageNotes.Count == 0 || _pg == null) return;
        var save = _colShift; _colShift = 0;
        double y0 = _pageBottom - _noteTotal;
        _pg.Prims.Add(new LinePrim { X1 = _pg.MarL, X2 = _pg.MarL + 144, Y1 = y0 + 3, Y2 = y0 + 3, Width = 0.5 });
        EmitBlocks(_pageNotes.SelectMany(NoteBlocks).ToList(), y0 + NoteSep);
        _colShift = save; _pageNotes.Clear(); _noteTotal = 0;
    }

    /// <summary>Riga di tabella più alta di una pagina: il contenuto di ogni cella prosegue a fette sulle pagine successive.</summary>
    void SplitRow(TableL t, int r, int hdr, double hdrH)
    {
        var row = t.Rows[r];
        int n = row.Cells.Count;
        var bi = new int[n]; var li = new int[n];
        while (true)
        {
            double top = _y;
            var ends = new double[n]; var marks = new int[n]; bool more = false, moved = false;
            for (int k = 0; k < n; k++)
            {
                var c = row.Cells[k];
                double y = top + c.MarT, y0 = y; marks[k] = _pg.Prims.Count;
                while (bi[k] < c.Content.Count)
                {
                    if (c.Content[bi[k]] is ParaL p)
                    {
                        double ys = y + (li[k] == 0 ? p.Before : 0), yy = ys; int e = li[k];
                        while (e < p.Lines.Count && yy + p.Lines[e].H <= _bottom + 0.01) { yy += p.Lines[e].H; e++; }
                        if (e == li[k] && !_pageHasBody && y == y0) { yy += p.Lines[e].H; e++; }   // pagina vuota: almeno una riga
                        if (e > li[k]) { EmitPara(p, li[k], e, ys); moved = true; y = yy; }
                        li[k] = e;
                        if (e < p.Lines.Count) break;
                        y += p.After; bi[k]++; li[k] = 0;
                    }
                    else { var tb = (TableL)c.Content[bi[k]]; EmitTable(tb, 0, tb.Rows.Count, y); y += tb.Height; bi[k]++; moved = true; }
                }
                ends[k] = y + c.MarB;
                more |= bi[k] < c.Content.Count;
            }
            if (!moved && more) { NewPage(true, false, false); if (hdr > 0 && r >= hdr && hdrH < _bottom - _top) { EmitTable(t, 0, hdr, _y); _y += hdrH; } continue; }
            double sh = more ? _bottom - top : ends.Max() - top;
            for (int k = n - 1; k >= 0; k--)
            {
                var c = row.Cells[k];
                if (c.Fill != null) _pg.Prims.Insert(marks[k], new RectPrim { X = c.X + _colShift, Y = top, W = c.W, H = sh, Fill = c.Fill });
                void Edge(Bdr? b, double x1, double y1, double x2, double y2) { if (b != null) _pg.Prims.Add(new LinePrim { X1 = x1 + _colShift, Y1 = y1, X2 = x2 + _colShift, Y2 = y2, Width = b.W, Color = b.Color }); }
                Edge(c.B[0], c.X, top, c.X + c.W, top); Edge(c.B[2], c.X, top + sh, c.X + c.W, top + sh);
                Edge(c.B[1], c.X, top, c.X, top + sh); Edge(c.B[3], c.X + c.W, top, c.X + c.W, top + sh);
            }
            _pageHasBody = true; _y = top + sh;
            if (!more) return;
            NewPage(true, false, false);
            if (hdr > 0 && r >= hdr && hdrH < _bottom - _top) { EmitTable(t, 0, hdr, _y); _y += hdrH; }
        }
    }

    // ================= emissione =================
    static readonly Dictionary<string, string> Hl = new(StringComparer.OrdinalIgnoreCase)
    {
        ["yellow"] = "FFFF00", ["green"] = "00FF00", ["cyan"] = "00FFFF", ["magenta"] = "FF00FF", ["blue"] = "0000FF", ["red"] = "FF0000",
        ["darkBlue"] = "000080", ["darkCyan"] = "008080", ["darkGreen"] = "008000", ["darkMagenta"] = "800080", ["darkRed"] = "800000",
        ["darkYellow"] = "808000", ["darkGray"] = "808080", ["lightGray"] = "C0C0C0", ["black"] = "000000", ["white"] = "FFFFFF"
    };

    void EmitPara(ParaL p, int from, int to, double y)
    {
        double paraTopY = y;
        if (from == 0 && p.Floats.Count > 0) EmitFloats(p, y, true);
        for (int li = from; li < to; li++)
        {
            var l = p.Lines[li];
            double baseY = y + l.Base;
            var lb = new LineBox { Para = p.P, PageIndex = _pg.Index, Start = l.Start, End = l.End, Story = p.Story, Y = p.Collapsed ? _lastLineY : y, Height = p.Collapsed ? l.NatH : l.H, Baseline = baseY, Left = p.X + _colShift + l.Left, Right = p.X + _colShift + l.Right, LastOfPara = l.Last };
            foreach (var s in l.Segs)
            {
                var a = s.A; double x = p.X + _colShift + s.X;
                var f = a.Fmt;
                double sh = f.VertAlign > 0 ? -f.Size * 0.35 : f.VertAlign < 0 ? f.Size * 0.14 : 0;
                lb.Segs.Add(new SegBox { Start = a.Start, Length = a.Len, X = x, Width = s.W, Text = a.Text, Fmt = f, Kind = a.K switch { 't' => LeafKind.Tab, 'b' => LeafKind.Break, 'p' => LeafKind.PageBreak, 'i' => LeafKind.Image, _ => LeafKind.Text }, Synthetic = a.Synth });
                if (a.K == 'i')
                {
                    _pg.Prims.Add(new ImagePrim { X = x, Y = baseY - a.ImgH, W = s.W, H = a.ImgH, Part = a.Img });
                    continue;
                }
                if (a.K is 't' or 'b' or 'p') { if (a.K == 't' && f.Underline) { } continue; }
                if (f.Highlight != null && Hl.TryGetValue(f.Highlight, out var hc))
                {
                    double lh = FontCatalog.LineHeight(f.Font, f.EffSize, f.Bold, f.Italic);
                    _pg.Prims.Add(new RectPrim { X = x, Y = baseY - FontCatalog.AscentOf(f.Font, f.EffSize, f.Bold, f.Italic) + sh, W = s.W, H = lh, Fill = hc });
                }
                if (a.K == 'w' || f.Underline || f.Strike)
                {
                    if (a.K == 'w') _pg.Prims.Add(new TextPrim { X = x, Baseline = baseY + sh, Text = a.Text, Fmt = f, Width = s.W });
                    if (f.Underline) _pg.Prims.Add(new LinePrim { X1 = x, X2 = x + s.W, Y1 = baseY + sh + f.EffSize * 0.12, Y2 = baseY + sh + f.EffSize * 0.12, Width = Math.Max(0.5, f.EffSize / 18), Color = f.Color ?? "000000" });
                    if (f.Strike) _pg.Prims.Add(new LinePrim { X1 = x, X2 = x + s.W, Y1 = baseY + sh - f.EffSize * 0.27, Y2 = baseY + sh - f.EffSize * 0.27, Width = Math.Max(0.5, f.EffSize / 18), Color = f.Color ?? "000000" });
                }
            }
            if (!p.Collapsed) _lastLineY = y;
            _pg.Lines.Add(lb); _res.Register(lb);
            y += l.H;
        }
        if (to == p.Lines.Count && p.Floats.Count > 0) EmitFloats(p, paraTopY, false);
    }

    void EmitBlocks(List<Block> bs, double y)
    {
        ParaL? prev = null;
        foreach (var b in bs)
        {
            if (b is ParaL p) { y += Gap(prev, p); EmitPara(p, 0, p.Lines.Count, y); y += p.Lines.Sum(l => l.H) + p.After; prev = p; }
            else if (b is TableL t) { EmitTable(t, 0, t.Rows.Count, y); y += t.Height; prev = null; }
        }
    }

    void EmitTable(TableL t, int r0, int r1, double y)
    {
        double yy = y;
        var rowY = new double[t.Rows.Count + 1];
        for (int r = r0; r < r1; r++) { rowY[r] = yy; yy += t.Rows[r].H; }
        rowY[r1] = yy;
        for (int r = r0; r < r1; r++)
            foreach (var c in t.Rows[r].Cells)
            {
                if (c.Continue) continue;
                int endRow = Math.Min(r + c.RowSpan, r1);
                double cy = rowY[r], ch = rowY[endRow] - cy;
                if (c.Fill != null) _pg.Prims.Add(new RectPrim { X = c.X + _colShift, Y = cy, W = c.W, H = ch, Fill = c.Fill });
                double off = c.VAlign == "center" ? Math.Max(0, (ch - c.MarT - c.MarB - c.ContentH) / 2) : c.VAlign == "bottom" ? Math.Max(0, ch - c.MarB - c.MarT - c.ContentH) : 0;
                EmitBlocks(c.Content, cy + c.MarT + off);
                void Edge(Bdr? b, double x1, double y1, double x2, double y2) { if (b != null) _pg.Prims.Add(new LinePrim { X1 = x1 + _colShift, Y1 = y1, X2 = x2 + _colShift, Y2 = y2, Width = b.W, Color = b.Color }); }
                Edge(c.B[0], c.X, cy, c.X + c.W, cy);
                Edge(c.B[2], c.X, cy + ch, c.X + c.W, cy + ch);
                Edge(c.B[1], c.X, cy, c.X, cy + ch);
                Edge(c.B[3], c.X + c.W, cy, c.X + c.W, cy + ch);
            }
    }

    // ================= costruzione paragrafi =================
    static readonly char[] PuaSym = [];

    static string FormatNum(int n, string fmt)
    {
        switch (fmt)
        {
            case "decimalZero": return n.ToString("00", CultureInfo.InvariantCulture);
            case "upperRoman": return Roman(n);
            case "lowerRoman": return Roman(n).ToLowerInvariant();
            case "upperLetter": return Letters(n);
            case "lowerLetter": return Letters(n).ToLowerInvariant();
            case "none": return "";
            default: return n.ToString(CultureInfo.InvariantCulture);
        }
    }
    static string Letters(int n) { var s = ""; n = Math.Max(1, n); while (n > 0) { n--; s = (char)('A' + n % 26) + s; n /= 26; } return s; }
    static string Roman(int n)
    {
        var t = new (int, string)[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") };
        var sb = new System.Text.StringBuilder();
        foreach (var (v, s) in t) while (n >= v) { sb.Append(s); n -= v; }
        return sb.ToString();
    }

    string NumberLabel(Ctx c, int numId, int ilvl, NumLevel lvl)
    {
        if (lvl.Fmt == "bullet")
        {
            var t = lvl.Text;
            if (t.Length > 0 && t[0] >= '') return t[0] == '' || lvl.Font.Equals("Symbol", StringComparison.OrdinalIgnoreCase) ? "•" : t[0] == '' ? "▪" : "•";
            return t;
        }
        if (!c.Counters.TryGetValue(numId, out var cs)) c.Counters[numId] = cs = new int[9];
        cs[ilvl] = cs[ilvl] == 0 ? lvl.Start : cs[ilvl] + 1;
        for (int j = ilvl + 1; j < 9; j++) cs[j] = 0;
        var text = lvl.Text;
        for (int k = 1; k <= 9; k++)
        {
            if (!text.Contains('%' + k.ToString())) continue;
            var lk = c.St.GetLevel(numId, k - 1);
            int val = cs[k - 1] == 0 ? (lk?.Start ?? 1) : cs[k - 1];
            text = text.Replace("%" + k, FormatNum(val, lk?.Fmt ?? "decimal"));
        }
        return text;
    }

    static string FixSymbolText(string t) => t;

    ParaL BuildPara(Paragraph p, double x, double width, string? tableStyle, Ctx c)
    {
        var F = c.St.ResolvePara(p, tableStyle, out var mark, out var baseRun);
        if (c.CondRun != null) { c.CondRun(mark); c.CondRun(baseRun); }
        var para = new ParaL { P = p, F = F, X = x, Width = width, Before = F.Before, After = F.After };
        var atoms = new List<Atom>();
        double indL = F.IndLeft, indR = F.IndRight, first = F.FirstLine;
        // etichetta di elenco
        if (F.NumId is > 0 && c.St.GetLevel(F.NumId.Value, F.Ilvl ?? 0) is { } lvl && lvl.Fmt != "none")
        {
            var lf = mark.Clone(); lvl.ApplyRun?.Invoke(lf);
            string lab = NumberLabel(c, F.NumId.Value, F.Ilvl ?? 0, lvl);
            if (lvl.Fmt == "bullet" && lab == "•" || lvl.Font.Equals("Symbol", StringComparison.OrdinalIgnoreCase) || lvl.Font.Equals("Wingdings", StringComparison.OrdinalIgnoreCase)) lf.Font = mark.Font;
            atoms.Add(new Atom { K = 'w', Text = lab, Fmt = lf, W = Metrics.Width(lab, lf), Synth = true, MetricFont = lvl.Font.Length > 0 && lvl.Fmt == "bullet" ? lvl.Font : null });
            if (lvl.Suff == "tab") atoms.Add(new Atom { K = 't', Text = "\t", Fmt = lf, Synth = true });
            else if (lvl.Suff == "space") atoms.Add(new Atom { K = 's', Text = " ", Fmt = lf, W = Metrics.Width(" ", lf), Synth = true, Spaces = 1 });
        }
        // richiami di nota: non sono foglie (niente offset), si inseriscono come atomi sintetici nell'ordine del documento
        var nrefs = new Queue<(Run Run, long Id)>(); Dictionary<Run, int>? ord = null;
        foreach (var r in p.Descendants<Run>())
        {
            long? nid = r.GetFirstChild<FootnoteReference>()?.Id?.Value ?? (r.GetFirstChild<FootnoteReferenceMark>() != null ? -1 : null);
            if (nid != null && r.Ancestors<Paragraph>().First() == p && !r.Ancestors<DeletedRun>().Any() && !r.Ancestors<AlternateContentFallback>().Any()) nrefs.Enqueue((r, nid.Value));
        }
        if (nrefs.Count > 0) ord = p.Descendants<Run>().Select((r, i) => (r, i)).ToDictionary(x => x.r, x => x.i);
        void NotesBefore(Run? run)
        {
            while (nrefs.Count > 0 && (run == null || ord![nrefs.Peek().Run] < ord[run])) { var (nr, id) = nrefs.Dequeue(); AddNoteMark(atoms, nr, id, baseRun, c); }
        }
        foreach (var leaf in ParaText.Leaves(p))
        {
            NotesBefore(leaf.Run);
            if (!c.RunCache.TryGetValue(leaf.Run, out var f)) c.RunCache[leaf.Run] = f = c.St.ResolveRun(baseRun, leaf.Run);
            switch (leaf.Kind)
            {
                case LeafKind.Tab: atoms.Add(new Atom { K = 't', Text = "\t", Fmt = f, Start = leaf.Start, Len = 1 }); break;
                case LeafKind.Break: atoms.Add(new Atom { K = 'b', Text = "\n", Fmt = f, Start = leaf.Start, Len = 1 }); break;
                case LeafKind.PageBreak: atoms.Add(new Atom { K = 'p', Text = "\n", Fmt = f, Start = leaf.Start, Len = 1 }); break;
                case LeafKind.Image:
                    {
                        var inl = leaf.El.GetFirstChild<DW.Inline>();
                        double w = (inl?.Extent?.Cx?.Value ?? 0) / 12700.0, h = (inl?.Extent?.Cy?.Value ?? 0) / 12700.0;
                        var rid = leaf.El.Descendants<A.Blip>().FirstOrDefault()?.Embed?.Value;
                        ImagePart? img = null;
                        try { if (rid != null) img = c.Part.GetPartById(rid) as ImagePart; } catch (ArgumentOutOfRangeException) { }
                        atoms.Add(new Atom { K = 'i', Text = "￼", Fmt = f, W = w, ImgH = h, Img = img, Start = leaf.Start, Len = 1 });
                        break;
                    }
                default:
                    {
                        var text = leaf.Text; var tf = f;
                        if (leaf.Field != null)
                        {
                            var fn = leaf.Field.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.ToUpperInvariant();
                            if (fn == "PAGE") text = c.PageNo.ToString(); else if (fn == "NUMPAGES") text = c.NumPages.ToString();
                        }
                        if (f.AllCaps || f.SmallCaps) { text = text.ToUpperInvariant(); if (f.SmallCaps && !f.AllCaps) { tf = f.Clone(); tf.Size *= 0.8; } }
                        if (text.Length != leaf.Length)
                        {
                            atoms.Add(new Atom { K = 'w', Text = text, Fmt = tf, W = Metrics.Width(text, tf), Start = leaf.Start, Len = leaf.Length });
                            break;
                        }
                        Tokenize(text, leaf.Start, tf, atoms);
                        break;
                    }
            }
        }
        NotesBefore(null);
        para.Floats = ReadFloats(p, c);
        // testo "sopra e sotto": lo spazio dell'oggetto ancorato al paragrafo viene riservato prima del testo
        foreach (var fl in para.Floats.Where(f => f.Wrap == "topBottom" && f.VFrom == "paragraph")) para.Before += fl.H + 2 * fl.Dist;
        BuildLines(para, atoms, mark, indL, indR, first, c);
        // paragrafo che contiene solo un'interruzione di sezione (pagina successiva): Word non gli dedica una riga
        if (atoms.Count == 0 && p.ParagraphProperties?.SectionProperties is { } sp0 && sp0.GetFirstChild<SectionType>()?.Val?.Value != SectionMarkValues.Continuous)
        { para.Collapsed = true; para.Before = para.After = 0; foreach (var ln in para.Lines) ln.H = 0; }
        para.Height = para.Before + para.Lines.Sum(l => l.H) + para.After;
        return para;
    }

    List<FloatL> ReadFloats(Paragraph p, Ctx c)
    {
        var res = new List<FloatL>();
        foreach (var anchor in p.Descendants().Where(e => e.LocalName == "anchor" && !e.Ancestors<AlternateContentFallback>().Any() && e.Ancestors<Paragraph>().First() == p))
        {
            double Emu(string? v) => double.TryParse(v, CultureInfo.InvariantCulture, out var d) ? d / 12700.0 : 0;
            var ext = anchor.ChildElements.FirstOrDefault(e => e.LocalName == "extent");
            var f = new FloatL { W = Emu(ext == null ? null : Attr(ext, "cx")), H = Emu(ext == null ? null : Attr(ext, "cy")), Behind = Attr(anchor, "behindDoc") is "1" or "true", Dist = Emu(Attr(anchor, "distT")) };
            if (f.W <= 0 || f.H <= 0) continue;
            var ph = anchor.ChildElements.FirstOrDefault(e => e.LocalName == "positionH"); var pv = anchor.ChildElements.FirstOrDefault(e => e.LocalName == "positionV");
            if (ph != null) { f.HFrom = Attr(ph, "relativeFrom") ?? "column"; f.HAlign = ph.ChildElements.FirstOrDefault(e => e.LocalName == "align")?.InnerText; f.OffX = Emu(ph.ChildElements.FirstOrDefault(e => e.LocalName == "posOffset")?.InnerText); }
            if (pv != null) { f.VFrom = Attr(pv, "relativeFrom") ?? "paragraph"; f.VAlign = pv.ChildElements.FirstOrDefault(e => e.LocalName == "align")?.InnerText; f.OffY = Emu(pv.ChildElements.FirstOrDefault(e => e.LocalName == "posOffset")?.InnerText); }
            var wrap = anchor.ChildElements.FirstOrDefault(e => e.LocalName.StartsWith("wrap", StringComparison.Ordinal))?.LocalName;
            f.Wrap = wrap switch { "wrapTopAndBottom" => "topBottom", "wrapSquare" or "wrapTight" or "wrapThrough" => "square", _ => "none" };
            var blip = anchor.Descendants<A.Blip>().FirstOrDefault();
            if (blip?.Embed?.Value is { } rid) { try { f.Img = c.Part.GetPartById(rid) as ImagePart; } catch (ArgumentOutOfRangeException) { } }
            var tx = anchor.Descendants().FirstOrDefault(e => e.LocalName == "txbxContent");
            if (tx != null) f.BoxParas = tx.Elements<Paragraph>().ToList();
            var fillClr = anchor.Descendants().FirstOrDefault(e => e.LocalName == "solidFill" && e.Parent?.LocalName == "spPr")?.Descendants().FirstOrDefault(e => e.LocalName == "srgbClr");
            f.Fill = fillClr == null ? null : Attr(fillClr, "val");
            var ln = anchor.Descendants().FirstOrDefault(e => e.LocalName == "ln" && e.Parent?.LocalName == "spPr");
            if (ln != null) { f.LineW = Emu(Attr(ln, "w")); f.LineColor = ln.Descendants().FirstOrDefault(e => e.LocalName == "srgbClr") is { } lc ? Attr(lc, "val") : null; if (f.LineColor != null && f.LineW <= 0) f.LineW = 0.75; }
            if (f.Img != null || f.BoxParas != null) res.Add(f);
        }
        return res;
    }

    void EmitFloats(ParaL p, double paraTop, bool behind)
    {
        var pg = _pg;
        foreach (var f in p.Floats.Where(f => f.Behind == behind))
        {
            double areaX = f.HFrom is "page" ? 0 : f.HFrom is "margin" or "leftMargin" or "insideMargin" ? pg.MarL : p.X + _colShift;
            double areaW = f.HFrom is "page" ? pg.Width : f.HFrom is "margin" ? pg.Width - pg.MarL - pg.MarR : p.Width;
            double x = f.HAlign switch { "center" => areaX + (areaW - f.W) / 2, "right" => areaX + areaW - f.W, "left" => areaX, _ => areaX + f.OffX };
            double areaY = f.VFrom is "page" or "topMargin" ? 0 : f.VFrom is "margin" ? pg.MarT : paraTop;
            double areaH = f.VFrom is "page" ? pg.Height : f.VFrom is "margin" ? pg.Height - pg.MarT - pg.MarB : f.H;
            double y = f.VAlign switch { "center" => areaY + (areaH - f.H) / 2, "bottom" => areaY + areaH - f.H, "top" => areaY, _ => areaY + f.OffY };
            if (f.Img != null) _pg.Prims.Add(new ImagePrim { X = x, Y = y, W = f.W, H = f.H, Part = f.Img });
            if (f.BoxParas != null)
            {
                if (f.Fill != null) _pg.Prims.Add(new RectPrim { X = x, Y = y, W = f.W, H = f.H, Fill = f.Fill });
                if (f.LineColor != null) foreach (var (x1, y1, x2, y2) in new[] { (x, y, x + f.W, y), (x, y + f.H, x + f.W, y + f.H), (x, y, x, y + f.H), (x + f.W, y, x + f.W, y + f.H) })
                    _pg.Prims.Add(new LinePrim { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Width = Math.Max(0.5, f.LineW), Color = f.LineColor });
                const double inset = 5.4;
                var ctx = new Ctx { St = _st, Part = _body.Part };
                var blocks = f.BoxParas.Select(bp => (Block)BuildPara(bp, x + inset, Math.Max(10, f.W - 2 * inset), null, ctx)).ToList();
                var save = _colShift; _colShift = 0;
                EmitBlocks(blocks, y + inset);
                _colShift = save;
            }
        }
    }

    static void Tokenize(string text, int start, RunFmt f, List<Atom> atoms)
    {
        int i = 0;
        while (i < text.Length)
        {
            int j = i;
            if (text[i] == ' ')
            {
                while (j < text.Length && text[j] == ' ') j++;
                var s = text[i..j];
                atoms.Add(new Atom { K = 's', Text = s, Fmt = f, W = Metrics.Width(s, f), Start = start + i, Len = j - i, Spaces = j - i });
            }
            else
            {
                while (j < text.Length && text[j] != ' ')
                {
                    j++;
                    if ((text[j - 1] == '-' || text[j - 1] == '–') && j < text.Length && text[j] != ' ' && j - 1 > i) break;
                }
                var s = text[i..j];
                atoms.Add(new Atom { K = 'w', Text = s, Fmt = f, W = Metrics.Width(s, f), Start = start + i, Len = j - i });
            }
            i = j;
        }
    }

    double NextTab(ParaFmt F, double x, double indL, StyleResolver st, out char kind, out string leader)
    {
        kind = 'l'; leader = "none";
        double best = double.MaxValue; TabDef? bt = null;
        foreach (var t in F.Tabs) if (t.Pos > x + 0.01 && t.Pos < best) { best = t.Pos; bt = t; }
        double lastCustom = F.Tabs.Count == 0 ? -1 : F.Tabs.Max(t => t.Pos);
        if (F.FirstLine < 0 && indL > x + 0.01 && indL < best) { best = indL; bt = null; }
        else if (x >= lastCustom)
        {
            double d = (Math.Floor(x / st.DefaultTab + 1e-9) + 1) * st.DefaultTab;
            if (d < best) { best = d; bt = null; }
        }
        if (bt != null)
        {
            kind = bt.Kind switch { "center" => 'c', "right" or "end" => 'r', "decimal" => 'd', _ => 'l' };
            leader = bt.Leader;
        }
        return best;
    }

    void BuildLines(ParaL para, List<Atom> atoms, RunFmt mark, double indL, double indR, double first, Ctx c)
    {
        var F = para.F;
        double right = para.Width - indR;
        LineL cur = new() { Left = indL + first };
        double x = cur.Left;
        bool content = false;
        double maxAsc = 0, maxDesc = 0, fontNat = 0;
        int lineStartOffset = 0;
        var lines = para.Lines;

        void Metr(RunFmt f, Atom? a = null)
        {
            if (a is { K: 'i' }) { maxAsc = Math.Max(maxAsc, a.ImgH); return; }
            var ff = FontCatalog.Get(a?.MetricFont ?? f.Font, f.Bold, f.Italic);
            maxAsc = Math.Max(maxAsc, (ff.Ascent + ff.LineGap) * f.EffSize);
            maxDesc = Math.Max(maxDesc, ff.Descent * f.EffSize);
            fontNat = Math.Max(fontNat, ff.LineHeightEm * f.EffSize);
        }

        void Finish(bool last, bool forced)
        {
            if (last) Metr(mark);
            if (maxAsc + maxDesc == 0) Metr(mark);
            double nat = maxAsc + maxDesc, h = nat, bl = maxAsc;
            if (F.LineRule == "exact") { h = F.Line / 20.0; bl = h * maxAsc / nat; }
            else if (F.LineRule == "atLeast") { h = Math.Max(nat, F.Line / 20.0); bl = maxAsc + (h - nat); }
            else if (F.Line != 240) { h = nat + fontNat * (F.Line / 240.0 - 1); bl = maxAsc + (h - nat); }   // il multiplo scala solo la parte di testo: le immagini inline no
            cur.H = h; cur.NatH = h; cur.Base = bl; cur.Last = last;
            // larghezza utilizzata senza spazi finali
            int endIdx = cur.Segs.Count;
            while (endIdx > 0 && cur.Segs[endIdx - 1].A.K == 's') endIdx--;
            double usedEnd = endIdx > 0 ? cur.Segs[endIdx - 1].X + cur.Segs[endIdx - 1].W : cur.Left;
            double avail = right - cur.Left;
            double slack = right - usedEnd;
            if (F.Alignment == Align.Justify && !last && !forced && slack > 0)
            {
                int sp = 0; for (int i = 0; i < endIdx; i++) sp += cur.Segs[i].A.K == 's' ? cur.Segs[i].A.Spaces : 0;
                if (sp > 0)
                {
                    double per = slack / sp, shift = 0;
                    for (int i = 0; i < cur.Segs.Count; i++)
                    {
                        cur.Segs[i].X += shift;
                        if (i < endIdx && cur.Segs[i].A.K == 's') { cur.Segs[i].W += per * cur.Segs[i].A.Spaces; shift += per * cur.Segs[i].A.Spaces; }
                    }
                    usedEnd = right;
                }
            }
            else if (F.Alignment is Align.Center or Align.Right && slack > 0)
            {
                double sh = F.Alignment == Align.Center ? slack / 2 : slack;
                foreach (var s in cur.Segs) s.X += sh;
            }
            cur.Right = right;
            cur.Start = cur.Segs.Where(s => !s.A.Synth).Select(s => s.A.Start).DefaultIfEmpty(lineStartOffset).Min();
            cur.End = cur.Segs.Where(s => !s.A.Synth).Select(s => s.A.Start + s.A.Len).DefaultIfEmpty(lineStartOffset).Max();
            if (cur.Segs.All(s => s.A.Synth)) { cur.Start = cur.End = lineStartOffset; }
            lineStartOffset = cur.End;
            lines.Add(cur);
            cur = new() { Left = indL };
            x = cur.Left; content = false; maxAsc = maxDesc = fontNat = 0;
        }

        for (int ai = 0; ai < atoms.Count; ai++)
        {
            var a = atoms[ai];
            switch (a.K)
            {
                case 's':
                    cur.Segs.Add(new SegL { A = a, X = x, W = a.W }); x += a.W; Metr(a.Fmt);
                    break;
                case 'w':
                case 'i':
                    {
                        if (x + a.W > right + 0.01 && content) { Finish(false, false); }
                        if (a.K == 'w' && x + a.W > right + 0.01 && !content && a.Text.Length > 1 && !a.Synth)
                        {
                            // parola più lunga della riga: spezza per caratteri
                            int cut = 1;
                            while (cut < a.Text.Length && x + Metrics.Width(a.Text[..(cut + 1)], a.Fmt) <= right + 0.01) cut++;
                            var head = new Atom { K = 'w', Text = a.Text[..cut], Fmt = a.Fmt, Start = a.Start, Len = cut, W = Metrics.Width(a.Text[..cut], a.Fmt) };
                            var tail = new Atom { K = 'w', Text = a.Text[cut..], Fmt = a.Fmt, Start = a.Start + cut, Len = a.Len - cut, W = Metrics.Width(a.Text[cut..], a.Fmt) };
                            atoms[ai] = tail; atoms.Insert(ai, head); a = head;
                        }
                        cur.Segs.Add(new SegL { A = a, X = x, W = a.W }); x += a.W; content = true; Metr(a.Fmt, a);
                        break;
                    }
                case 't':
                    {
                        double stop = NextTab(F, x, indL, c.St, out var kind, out var leader);
                        if (stop > right) stop = right;
                        double nx = stop;
                        if (kind != 'l')
                        {
                            double follow = 0;
                            for (int j = ai + 1; j < atoms.Count && atoms[j].K is 'w' or 's' or 'i'; j++)
                            {
                                if (kind == 'd' && atoms[j].K == 'w') { int d = atoms[j].Text.IndexOfAny(['.', ',']); follow += d >= 0 ? Metrics.Width(atoms[j].Text[..d], atoms[j].Fmt) : atoms[j].W; if (d >= 0) break; }
                                else follow += atoms[j].W;
                            }
                            nx = kind == 'c' ? stop - follow / 2 : stop - follow;
                            if (nx < x) nx = x;
                        }
                        a.W = nx - x;
                        cur.Segs.Add(new SegL { A = a, X = x, W = a.W }); x = nx; content = true; Metr(a.Fmt);
                        break;
                    }
                case 'b':
                case 'p':
                    cur.Segs.Add(new SegL { A = a, X = x, W = 0 }); Metr(a.Fmt);
                    cur.PageBreakAfter = a.K == 'p';
                    Finish(false, true);
                    // una riga vuota finale dopo il break è la riga del segno di paragrafo
                    if (ai == atoms.Count - 1 && a.K == 'b') { Finish(true, true); return; }
                    if (ai == atoms.Count - 1 && a.K == 'p') { lines[^1].Last = true; return; }
                    break;
            }
        }
        Finish(true, true);
    }

    // ================= tabelle =================
    static string? Attr(OpenXmlElement e, string name) => e.GetAttributes().FirstOrDefault(a => a.LocalName == name).Value;

    Bdr? ReadBdr(OpenXmlElement? e)
    {
        if (e == null) return null;
        var v = Attr(e, "val");
        if (v is null or "nil" or "none") return null;
        double sz = double.TryParse(Attr(e, "sz"), CultureInfo.InvariantCulture, out var s) ? s / 8.0 : 0.5;
        var col = Attr(e, "color");
        return new Bdr(Math.Max(0.5, sz), col is null or "auto" ? "000000" : col);
    }

    static OpenXmlElement? Child(OpenXmlElement? parent, params string[] names) => parent?.ChildElements.FirstOrDefault(c => names.Contains(c.LocalName));

    TableL BuildTable(Table tbl, double x, double width, Ctx c)
    {
        var t = new TableL();
        var tp = tbl.GetFirstChild<TableProperties>();
        var sid = tp?.TableStyle?.Val?.Value;
        // bordi e margini: stile tabella poi diretto
        var borders = new Bdr?[6];   // top,left,bottom,right,insideH,insideV
        double marL = 5.4, marR = 5.4, marT = 0, marB = 0;
        void ApplyBorders(OpenXmlElement? tb)
        {
            if (tb == null) return;
            string[][] names = [["top"], ["left", "start"], ["bottom"], ["right", "end"], ["insideH"], ["insideV"]];
            for (int i = 0; i < 6; i++) { var e = Child(tb, names[i]); if (e != null) borders[i] = ReadBdr(e); }
        }
        void ApplyMar(OpenXmlElement? mar)
        {
            if (mar == null) return;
            double? V(params string[] n) => Child(mar, n) is { } e && double.TryParse(Attr(e, "w"), CultureInfo.InvariantCulture, out var d) ? d / 20.0 : null;
            marL = V("left", "start") ?? marL; marR = V("right", "end") ?? marR; marT = V("top") ?? marT; marB = V("bottom") ?? marB;
        }
        foreach (var s in c.St.Chain(sid ?? c.St.DefaultTableStyleId)) { var stp = s.GetFirstChild<StyleTableProperties>(); ApplyBorders(Child(stp, "tblBorders")); ApplyMar(Child(stp, "tblCellMar")); }
        ApplyBorders(Child(tp, "tblBorders")); ApplyMar(Child(tp, "tblCellMar"));
        double ind = Child(tp, "tblInd") is { } ti && double.TryParse(Attr(ti, "w"), CultureInfo.InvariantCulture, out var iv) ? iv / 20.0 : 0;
        t.X = x + ind - (_compat15 ? 0 : marL);

        var grid = tbl.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Select(g => (double.TryParse(g.Width?.Value, CultureInfo.InvariantCulture, out var gw) ? gw : 0) / 20.0).ToList() ?? new();
        var trs = tbl.Elements<TableRow>().ToList();
        if (grid.Count == 0 && trs.Count > 0)
            foreach (var tc in trs[0].Elements<TableCell>())
            {
                var w = tc.TableCellProperties?.TableCellWidth?.Width?.Value;
                int span = tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                for (int k = 0; k < span; k++) grid.Add((double.TryParse(w, out var wd) ? wd / 20.0 : 72) / span);
            }
        int need = trs.Select(r => r.Elements<TableCell>().Sum(cc => cc.TableCellProperties?.GridSpan?.Val?.Value ?? 1) + (r.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0)).DefaultIfEmpty(0).Max();
        while (grid.Count < need) grid.Add(grid.Count > 0 ? grid[^1] : 72);
        t.Grid = grid.ToArray();
        var colX = new double[grid.Count + 1];
        colX[0] = t.X + (_compat15 ? 0 : 0);
        for (int i = 0; i < grid.Count; i++) colX[i + 1] = colX[i] + grid[i];

        // condizioni dello stile tabella (tblStylePr) e opzioni di aspetto (tblLook)
        var cond = new Dictionary<string, (string? Fill, OpenXmlElement? Brd, OpenXmlElement? RPr)>();
        foreach (var st in c.St.Chain(sid ?? c.St.DefaultTableStyleId))
            foreach (var sp in st.ChildElements.Where(x => x.LocalName == "tblStylePr"))
            {
                var type = Attr(sp, "type"); if (type == null) continue;
                var tcPr = Child(sp, "tcPr"); var shd = Child(tcPr, "shd"); var fl = shd != null ? Attr(shd, "fill") : null;
                cond.TryGetValue(type, out var old);
                cond[type] = (fl is null or "auto" ? old.Fill : fl, Child(tcPr, "tcBorders") ?? old.Brd, Child(sp, "rPr") ?? old.RPr);
            }
        var lk = tp?.ChildElements.FirstOrDefault(x => x.LocalName == "tblLook");
        int lv = lk != null && int.TryParse(Attr(lk, "val"), System.Globalization.NumberStyles.HexNumber, null, out var lvv) ? lvv : 0x04A0;
        bool Look(string name, int bit) => lk != null && Attr(lk, name) is { } a ? a is "1" or "true" : (lv & bit) != 0;
        bool lFirstRow = Look("firstRow", 0x20), lLastRow = Look("lastRow", 0x40), lFirstCol = Look("firstColumn", 0x80), lLastCol = Look("lastColumn", 0x100), lNoH = Look("noHBand", 0x200), lNoV = Look("noVBand", 0x400);
        var open = new Dictionary<int, CellL>();     // colonna → cella con unione verticale aperta
        int nrows = trs.Count;
        for (int ri = 0; ri < nrows; ri++)
        {
            var tr = trs[ri];
            var row = new RowL();
            int col = tr.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
            foreach (var tc in tr.Elements<TableCell>())
            {
                var cp = tc.TableCellProperties;
                int span = Math.Max(1, cp?.GridSpan?.Val?.Value ?? 1);
                var cell = new CellL { Col = col, Span = span, X = colX[Math.Min(col, grid.Count)], W = colX[Math.Min(col + span, grid.Count)] - colX[Math.Min(col, grid.Count)], MarT = marT, MarB = marB };
                double cml = marL, cmr = marR;
                if (Child(cp, "tcMar") is { } tm)
                {
                    double? V(params string[] n) => Child(tm, n) is { } e && double.TryParse(Attr(e, "w"), CultureInfo.InvariantCulture, out var d) ? d / 20.0 : null;
                    cml = V("left", "start") ?? cml; cmr = V("right", "end") ?? cmr; cell.MarT = V("top") ?? cell.MarT; cell.MarB = V("bottom") ?? cell.MarB;
                }
                var vm = cp?.VerticalMerge;
                if (vm != null)
                {
                    bool restart = vm.Val?.Value == MergedCellValues.Restart;
                    if (restart) open[col] = cell;
                    else if (open.TryGetValue(col, out var top)) { top.RowSpan++; cell.Continue = true; }
                }
                else open.Remove(col);
                // formattazione condizionale: dal più generico al più specifico
                var conds = new List<string> { "wholeTable" };
                int dataRow = ri - (lFirstRow ? 1 : 0);
                if (!lNoV) conds.Add(((col - (lFirstCol ? 1 : 0)) % 2 == 0 ? "band1Vert" : "band2Vert"));
                if (!lNoH && !(lFirstRow && ri == 0) && !(lLastRow && ri == nrows - 1)) conds.Add(dataRow % 2 == 0 ? "band1Horz" : "band2Horz");
                if (lLastCol && col + span >= grid.Count) conds.Add("lastCol");
                if (lFirstCol && col == 0) conds.Add("firstCol");
                if (lLastRow && ri == nrows - 1) conds.Add("lastRow");
                if (lFirstRow && ri == 0) conds.Add("firstRow");
                string? cFill = null; OpenXmlElement? cBrd = null; var cRun = new List<OpenXmlElement>();
                foreach (var cn in conds) if (cond.TryGetValue(cn, out var cv)) { cFill = cv.Fill ?? cFill; cBrd = cv.Brd ?? cBrd; if (cv.RPr != null) cRun.Add(cv.RPr); }
                cell.Fill = cFill;
                var condRun = cRun.Count == 0 ? null : (Action<RunFmt>)(f => { foreach (var rp in cRun) c.St.ApplyRPr(f, rp); });
                // bordi
                var tcb = Child(cp, "tcBorders");
                Bdr? Pick(int tblIdx, string[] n, bool edge) => tcb != null && Child(tcb, n) is { } e ? ReadBdr(e) : cBrd != null && Child(cBrd, edge ? n : [tblIdx == 4 ? "insideH" : "insideV"]) is { } e2 ? ReadBdr(e2) : borders[tblIdx];
                bool firstRow = ri == 0, lastRow = ri == nrows - 1, firstCol = col == 0, lastCol = col + span >= grid.Count;
                cell.B[0] = Pick(firstRow ? 0 : 4, ["top"], firstRow);
                cell.B[2] = Pick(lastRow ? 2 : 4, ["bottom"], lastRow);
                cell.B[1] = Pick(firstCol ? 1 : 5, ["left", "start"], firstCol);
                cell.B[3] = Pick(lastCol ? 3 : 5, ["right", "end"], lastCol);
                if (cp?.Shading?.Fill?.Value is { } fill && fill != "auto") cell.Fill = fill;
                c.CondRun = condRun;
                cell.VAlign = cp?.TableCellVerticalAlignment?.Val?.InnerText ?? "top";
                if (!cell.Continue)
                {
                    var inner = BuildCellBlocks(tc, cell.X + cml, Math.Max(10, cell.W - cml - cmr), sid, c);
                    cell.Content = inner;
                    cell.ContentH = StackH(inner);
                    c.CondRun = null;
                }
                cell.H = cell.ContentH + cell.MarT + cell.MarB;
                row.Cells.Add(cell);
                col += span;
            }
            row.Src = tr; row.IsHeader = tr.TableRowProperties?.GetFirstChild<TableHeader>() != null;
            t.Rows.Add(row);
        }
        // altezze di riga: dopo aver costruito tutte le righe (i RowSpan si conoscono solo a fine tabella)
        foreach (var row in t.Rows)
        {
            var tr = row.Src!;
            var trh = tr.TableRowProperties?.GetFirstChild<TableRowHeight>();
            double rh = row.Cells.Where(k => k.RowSpan == 1 && !k.Continue).Select(k => k.H).DefaultIfEmpty(0).Max();
            if (trh?.Val?.Value is { } hv)
            {
                bool exact = trh.HeightType?.Value == HeightRuleValues.Exact;
                rh = exact ? hv / 20.0 : Math.Max(rh, hv / 20.0);
            }
            if (rh == 0) rh = 14;
            row.H = rh;
        }
        // celle con unione verticale: se il contenuto eccede, allunga l'ultima riga del gruppo
        for (int ri = 0; ri < t.Rows.Count; ri++)
            foreach (var cell in t.Rows[ri].Cells.Where(k => k.RowSpan > 1))
            {
                int end = Math.Min(ri + cell.RowSpan - 1, t.Rows.Count - 1);
                double have = Enumerable.Range(ri, end - ri + 1).Sum(k => t.Rows[k].H);
                if (cell.H > have) t.Rows[end].H += cell.H - have;
            }
        t.Height = t.Rows.Sum(r => r.H);
        return t;
    }

    List<Block> BuildCellBlocks(TableCell tc, double x, double w, string? sid, Ctx c)
    {
        var list = new List<Block>();
        foreach (var e in Flatten(tc.ChildElements))
            if (e is Paragraph p) list.Add(BuildPara(p, x, w, sid, c));
            else if (e is Table t) list.Add(BuildTable(t, x, w, c));
        return list;
    }

    double StackH(List<Block> bs)
    {
        double h = 0; ParaL? prev = null;
        foreach (var b in bs)
        {
            if (b is ParaL p) { h += Gap(prev, p) + p.Lines.Sum(l => l.H) + p.After; prev = p; }
            else { h += b.Height; prev = null; }
        }
        return h;
    }
}
