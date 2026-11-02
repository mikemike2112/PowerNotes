using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Layout;
using PowerNotes.Core.Model;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

namespace PowerNotes.App;

/// <summary>Superficie di editing custom ("Layout di stampa"): disegna le pagine prodotte dal motore di layout, con caret e selezione. Non è un RichTextBox.</summary>
public sealed class DocumentView : FrameworkElement
{
    const double GapDip = 18, SideDip = 28;
    EditorSession? _s;
    LayoutResult? _cachedFor;
    readonly Dictionary<int, DrawingGroup> _cache = new();
    bool _caretOn = true;
    readonly DispatcherTimer _blink = new() { Interval = TimeSpan.FromMilliseconds(530) };
    bool _dragging;

    public ScrollViewer? Scroller { get; set; }
    /// <summary>Visualizza griglia: contorni tratteggiati delle celle di tabella (solo a schermo).</summary>
    public bool ShowTableGrid { get; set; }
    /// <summary>Griglia della pagina (solo a schermo): linee ogni 0,5 cm.</summary>
    public bool ShowGrid { get; set; }
    /// <summary>Mostra tutto: ¶ a fine paragrafo, punti sugli spazi, frecce sulle tabulazioni, ↵ sugli a capo (solo a schermo).</summary>
    public bool ShowMarks { get; set; }
    /// <summary>Modalità di lettura: il documento si legge e si copia ma non si modifica.</summary>
    public bool ReadOnly { get; set; }
    /// <summary>Pagine affiancate per riga (1 = una sola colonna).</summary>
    public int PagesPerRow { get; private set; } = 1;
    double _maxPw;
    public double Zoom { get; private set; } = 1.0;
    public event Action? StateChanged;
    /// <summary>Ctrl+clic su un collegamento: indirizzo (o «#segnalibro»).</summary>
    public event Action<string>? LinkClicked;
    public Brush? PaperShadow { get; set; } = new SolidColorBrush(Color.FromArgb(70, 0, 0, 0));

    double Scale => Zoom * 96.0 / 72.0;

    public DocumentView()
    {
        Focusable = true; FocusVisualStyle = null; Cursor = Cursors.IBeam;
        HorizontalAlignment = HorizontalAlignment.Stretch; VerticalAlignment = VerticalAlignment.Top;
        _blink.Tick += (_, _) => { _caretOn = !_caretOn; InvalidateVisual(); };
        _blink.Start();
        // IME / testo: ricevuto da OnTextInput
        InputMethod.SetIsInputMethodEnabled(this, true);
    }

    public EditorSession? Session
    {
        get => _s;
        set { _s = value; _cachedFor = null; Refresh(); }
    }

    public int PageCount => _s?.Layout.Pages.Count ?? 0;

    public int CaretPage
    {
        get { if (_s == null) return 0; var l = _s.Layout.Locate(_s.Caret.P, _s.Caret.Off); return (l?.PageIndex ?? 0) + 1; }
    }

    public void SetZoom(double z)
    {
        Zoom = Math.Clamp(z, 0.25, 4.0);
        InvalidateMeasure(); InvalidateVisual(); StateChanged?.Invoke();
    }

    /// <summary>Da chiamare dopo ogni modifica o cambio di selezione.</summary>
    public void Refresh(bool scrollToCaret = false)
    {
        InvalidateMeasure(); InvalidateVisual();
        _caretOn = true; _blink.Stop(); _blink.Start();
        if (scrollToCaret) Dispatcher.BeginInvoke(DispatcherPriority.Loaded, ScrollToCaret);
        StateChanged?.Invoke();
    }

    // ---- geometria ----
    (double w, double h, List<double> tops) Metrics()
    {
        var tops = new List<double>();
        double y = GapDip, w = 0;
        int n = Math.Max(1, PagesPerRow);
        if (_s != null)
        {
            var pages = _s.Layout.Pages;
            for (int i = 0; i < pages.Count; i += n)
            {
                double rowH = 0;
                for (int j = i; j < Math.Min(pages.Count, i + n); j++) { rowH = Math.Max(rowH, pages[j].Height * Scale); w = Math.Max(w, pages[j].Width * Scale); }
                for (int j = i; j < Math.Min(pages.Count, i + n); j++) tops.Add(y);
                y += rowH + GapDip;
            }
        }
        _maxPw = w;
        return (n * w + (n - 1) * GapDip + 2 * SideDip, y, tops);
    }

    protected override Size MeasureOverride(Size availableSize)
    {
        var (w, h, _) = Metrics();
        return new Size(w, h);
    }

    Rect PageRect(int i, List<double> tops)
    {
        var p = _s!.Layout.Pages[i];
        double pw = p.Width * Scale;
        int n = Math.Max(1, PagesPerRow);
        double rowW = n * _maxPw + (n - 1) * GapDip;
        double left = Math.Max(SideDip, (ActualWidth - rowW) / 2) + (i % n) * (_maxPw + GapDip) + (_maxPw - pw) / 2;
        return new Rect(left, tops[i], pw, p.Height * Scale);
    }

    // ---- disegno ----
    protected override void OnRender(DrawingContext dc)
    {
        dc.DrawRectangle(Brushes.Transparent, null, new Rect(0, 0, ActualWidth, ActualHeight));
        if (_s == null) return;
        var L = _s.Layout;
        if (!ReferenceEquals(L, _cachedFor)) { _cache.Clear(); _cachedFor = L; }
        var (_, _, tops) = Metrics();
        double vTop = Scroller?.VerticalOffset ?? 0, vBot = vTop + (Scroller?.ViewportHeight ?? ActualHeight);
        double ppd = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        var idx = new Dictionary<Paragraph, int>();
        var paras = _s.Paragraphs();
        for (int i = 0; i < paras.Count; i++) idx[paras[i]] = i;
        var (s, e) = _s.Range;
        int si = (idx.TryGetValue(s.P, out var q1) ? q1 : 0), ei = (idx.TryGetValue(e.P, out var q2) ? q2 : 0);
        var selBrush = new SolidColorBrush(Color.FromArgb(90, 11, 204, 234));
        LineBox? caretLine = _s.Layout.Locate(_s.Caret.P, _s.Caret.Off);

        for (int i = 0; i < L.Pages.Count; i++)
        {
            var r = PageRect(i, tops);
            if (r.Bottom < vTop - 50 || r.Top > vBot + 50) continue;
            dc.DrawRectangle(PaperShadow, null, new Rect(r.X + 3, r.Y + 3, r.Width, r.Height));
            dc.PushTransform(new TranslateTransform(r.X, r.Y));
            dc.PushTransform(new ScaleTransform(Scale, Scale));
            if (!_cache.TryGetValue(i, out var g))
            {
                g = new DrawingGroup();
                using (var gc = g.Open()) PageRenderer.Draw(gc, L.Pages[i], ppd);
                g.Freeze();
                _cache[i] = g;
            }
            dc.DrawDrawing(g);
            if (ShowGrid)
            {
                var gl = new Pen(new SolidColorBrush(Color.FromArgb(70, 90, 90, 90)), 0.4); gl.Freeze();
                const double Step = 28.3465 / 2;
                for (double gx = Step; gx < L.Pages[i].Width; gx += Step) dc.DrawLine(gl, new Point(gx, 0), new Point(gx, L.Pages[i].Height));
                for (double gy = Step; gy < L.Pages[i].Height; gy += Step) dc.DrawLine(gl, new Point(0, gy), new Point(L.Pages[i].Width, gy));
            }
            if (ShowMarks) DrawMarks(dc, L.Pages[i], ppd);
            if (ShowTableGrid && L.Pages[i].Cells.Count > 0)
            {
                var gp = new Pen(new SolidColorBrush(Color.FromArgb(160, 140, 140, 140)), 0.6) { DashStyle = DashStyles.Dash }; gp.Freeze();
                foreach (var (cx, cy, cw, ch) in L.Pages[i].Cells) dc.DrawRectangle(null, gp, new Rect(cx, cy, cw, ch));
            }
            if (_s.EditingHF) dc.DrawRectangle(new SolidColorBrush(Color.FromArgb(150, 255, 255, 255)), null, new Rect(0, L.Pages[i].MarT, L.Pages[i].Width, L.Pages[i].Height - L.Pages[i].MarT - L.Pages[i].MarB));
            if (_s.HasSelection)
                foreach (var l in L.Pages[i].Lines)
                {
                    if (!idx.TryGetValue(l.Para, out var ip) || ip < si || ip > ei) continue;
                    int a = ip == si ? Math.Max(s.Off, l.Start) : l.Start;
                    int b = ip == ei ? Math.Min(e.Off, l.End) : l.End;
                    double x1 = l.XAt(a), x2 = l.XAt(b);
                    if (ip < ei && l.LastOfPara) x2 += 4;      // segno di paragrafo selezionato
                    if (x2 <= x1 && !(a < b)) continue;
                    dc.DrawRectangle(selBrush, null, new Rect(x1, l.Y, Math.Max(1, x2 - x1), l.Height));
                }
            if (_caretOn && !_s.HasSelection && caretLine != null && caretLine.PageIndex == i && IsKeyboardFocused)
            {
                double cx = caretLine.XAt(_s.Caret.Off);
                dc.DrawRectangle(Brushes.Black, null, new Rect(cx, caretLine.Y, 0.9 / Scale * 1.0 + 0.4, caretLine.Height));
            }
            dc.Pop(); dc.Pop();
        }
    }

    static readonly Brush MarkBrush = Freeze(new SolidColorBrush(Color.FromRgb(0x5b, 0x8d, 0xc9)));
    static Brush Freeze(SolidColorBrush b) { b.Freeze(); return b; }

    static void Mark(DrawingContext dc, string text, double x, double baseline, double size, double ppd)
    {
        var ft = new FormattedText(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight, new Typeface("Segoe UI Symbol"), size, MarkBrush, ppd);
        dc.DrawText(ft, new Point(x, baseline - ft.Baseline));
    }

    void DrawMarks(DrawingContext dc, PowerNotes.Core.Layout.Page pg, double ppd)
    {
        foreach (var l in pg.Lines)
        {
            double size = Math.Clamp((l.Segs.FirstOrDefault(x => !x.Synthetic)?.Fmt.EffSize ?? 11) * 0.9, 6, 20), endX = l.Left;
            foreach (var sg in l.Segs)
            {
                if (sg.Synthetic) continue;
                endX = sg.X + sg.Width;
                if (sg.Kind == LeafKind.Tab) Mark(dc, "\u2192", sg.X + 1, l.Baseline, size, ppd);
                else if (sg.Kind == LeafKind.Break) Mark(dc, "\u21B5", sg.X, l.Baseline, size, ppd);
                else if (sg.Kind == LeafKind.Text && sg.Text.Length > 0 && sg.Text.All(c => c == ' '))
                {
                    double w = sg.Width / sg.Text.Length;
                    for (int k = 0; k < sg.Text.Length; k++) Mark(dc, "\u00B7", sg.X + k * w + w / 2 - size * 0.12, l.Baseline - size * 0.1, size, ppd);
                }
            }
            if (l.LastOfPara) Mark(dc, "\u00B6", endX + 1, l.Baseline, size, ppd);
        }
    }

    void ScrollToCaret()
    {
        if (_s == null || Scroller == null) return;
        var l = _s.Layout.Locate(_s.Caret.P, _s.Caret.Off);
        if (l == null) return;
        var (_, _, tops) = Metrics();
        var r = PageRect(l.PageIndex, tops);
        double y = r.Y + l.Y * Scale, h = l.Height * Scale;
        double top = Scroller.VerticalOffset, vp = Scroller.ViewportHeight;
        if (y < top + 8) Scroller.ScrollToVerticalOffset(Math.Max(0, y - 40));
        else if (y + h > top + vp - 8) Scroller.ScrollToVerticalOffset(y + h - vp + 40);
        double x = r.X + l.XAt(_s.Caret.Off) * Scale;
        if (x < Scroller.HorizontalOffset + 8) Scroller.ScrollToHorizontalOffset(Math.Max(0, x - 40));
        else if (x > Scroller.HorizontalOffset + Scroller.ViewportWidth - 8) Scroller.ScrollToHorizontalOffset(x - Scroller.ViewportWidth + 60);
    }

    // ---- mouse ----
    (int page, double x, double y)? ToPage(Point p)
    {
        if (_s == null || _s.Layout.Pages.Count == 0) return null;
        var (_, _, tops) = Metrics();
        int best = 0; double bd = double.MaxValue;
        for (int i = 0; i < tops.Count; i++)
        {
            var pr = PageRect(i, tops);
            double dx = p.X < pr.Left ? pr.Left - p.X : p.X > pr.Right ? p.X - pr.Right : 0, dy = p.Y < pr.Top ? pr.Top - p.Y : p.Y > pr.Bottom ? p.Y - pr.Bottom : 0;
            double d = dy * 1000 + dx;
            if (d < bd) { bd = d; best = i; }
        }
        var r = PageRect(best, tops);
        return (best, (p.X - r.X) / Scale, (p.Y - r.Y) / Scale);
    }

    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        Focus();
        bool started = ClickAt(e.GetPosition(this), Keyboard.Modifiers.HasFlag(ModifierKeys.Shift), e.ClickCount);
        _dragging = started && e.ClickCount == 1;
        if (started && e.ClickCount == 1 && Keyboard.Modifiers.HasFlag(ModifierKeys.Control) && _s?.HyperlinkAt() is { } link) { _dragging = false; LinkClicked?.Invoke(link); }
        if (_dragging) CaptureMouse();
        e.Handled = true;
    }

    /// <summary>Click in coordinate della vista (DIP): sposta caret/selezione. Usato dal mouse e dall'autotest.</summary>
    public bool ClickAt(Point pos, bool extend, int clickCount = 1)
    {
        if (_s == null) return false;
        var hit = ToPage(pos);
        if (hit == null) return false;
        var pg = _s.Layout.Pages[hit.Value.page];
        bool zoneTop = hit.Value.y < pg.MarT, zoneBottom = hit.Value.y > pg.Height - pg.MarB;
        if (clickCount == 2 && (zoneTop || zoneBottom) && !_s.EditingHF) { _s.EnterHeaderFooter(zoneTop); Refresh(true); return true; }
        if (clickCount == 2 && _s.EditingHF && !zoneTop && !zoneBottom) { _s.ExitHeaderFooter(); }
        _s.SetCaretAt(hit.Value.page, hit.Value.x, hit.Value.y, extend);
        if (clickCount == 2) _s.SelectWord();
        else if (clickCount >= 3) { _s.SetCaret(_s.Caret.P, 0); _s.SetCaret(_s.Caret.P, ParaText.Length(_s.Caret.P), true); }
        Refresh();
        return true;
    }

    /// <summary>Punto della vista (DIP) corrispondente a una posizione del documento.</summary>
    public Point? PointFor(Paragraph p, int off)
    {
        if (_s == null) return null;
        var l = _s.Layout.Locate(p, off);
        if (l == null) return null;
        var (_, _, tops) = Metrics();
        var r = PageRect(l.PageIndex, tops);
        return new Point(r.X + l.XAt(off) * Scale, r.Y + (l.Y + l.Height / 2) * Scale);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        if (!_dragging || _s == null) return;
        var hit = ToPage(e.GetPosition(this));
        if (hit == null) return;
        _s.SetCaretAt(hit.Value.page, hit.Value.x, hit.Value.y, true);
        Refresh(true);
    }

    protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e) { if (_dragging) { _dragging = false; ReleaseMouseCapture(); } }

    protected override void OnMouseWheel(MouseWheelEventArgs e)
    {
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) { SetZoom(Zoom * (e.Delta > 0 ? 1.1 : 1 / 1.1)); e.Handled = true; }
    }

    protected override void OnGotKeyboardFocus(KeyboardFocusChangedEventArgs e) => InvalidateVisual();
    protected override void OnLostKeyboardFocus(KeyboardFocusChangedEventArgs e) => InvalidateVisual();

    // ---- tastiera ----
    protected override void OnTextInput(TextCompositionEventArgs e)
    {
        if (_s == null || ReadOnly || string.IsNullOrEmpty(e.Text)) return;
        var t = new string(e.Text.Where(c => !char.IsControl(c) || c == '\t').ToArray());
        if (t.Length == 0) return;
        _s.Type(t); Refresh(true); e.Handled = true;
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (_s == null) return;
        bool shift = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift), ctrl = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
        bool edit = false, handled = true;
        var key = e.Key == Key.System ? e.SystemKey : e.Key;
        if (ReadOnly && key is Key.Back or Key.Delete or Key.Enter or Key.Tab) { e.Handled = true; return; }
        switch (key)
        {
            case Key.Left: _s.MoveHorizontal(-1, shift, ctrl); break;
            case Key.Right: _s.MoveHorizontal(1, shift, ctrl); break;
            case Key.Up: _s.MoveVertical(-1, shift); break;
            case Key.Down: _s.MoveVertical(1, shift); break;
            case Key.PageUp: _s.MoveVertical(-1, shift, ViewportPt()); break;
            case Key.PageDown: _s.MoveVertical(1, shift, ViewportPt()); break;
            case Key.Home: _s.MoveLineEdge(false, shift, ctrl); break;
            case Key.End: _s.MoveLineEdge(true, shift, ctrl); break;
            case Key.Escape: _s.ExitHeaderFooter(); break;
            case Key.Back: _s.Backspace(); edit = true; break;
            case Key.Delete: _s.DeleteForward(); edit = true; break;
            case Key.Enter: if (shift) _s.SoftBreak(); else _s.Enter(); edit = true; break;
            case Key.Tab:
                if (!_s.TabInTable(shift)) { if (_s.AtListStart) _s.ChangeIndent(shift ? -36 : 36); else _s.Type("\t"); }
                edit = true; break;
            default: handled = false; break;
        }
        if (handled) { e.Handled = true; Refresh(true); }
        _ = edit;
    }

    double ViewportPt() => (Scroller?.ViewportHeight ?? 600) / Scale * 0.9;

    /// <summary>Geometria (DIP, nelle coordinate della vista) della pagina del caret, per i righelli.</summary>
    public (Rect Rect, PowerNotes.Core.Layout.Page Page, double Scale)? CaretPageGeometry()
    {
        if (_s == null || _s.Layout.Pages.Count == 0) return null;
        var l = _s.Layout.Locate(_s.Caret.P, _s.Caret.Off);
        int i = Math.Clamp(l?.PageIndex ?? 0, 0, _s.Layout.Pages.Count - 1);
        var (_, _, tops) = Metrics();
        return (PageRect(i, tops), _s.Layout.Pages[i], Scale);
    }

    /// <summary>Affianca n pagine per riga (1 = una colonna) e porta lo zoom perché stiano tutte nella larghezza.</summary>
    public void FitPages(double viewportW, int n)
    {
        if (_s == null || _s.Layout.Pages.Count == 0) return;
        PagesPerRow = Math.Clamp(n, 1, 4);
        SetZoom((viewportW - 2 * SideDip - (PagesPerRow - 1) * GapDip) / (PagesPerRow * _s.Layout.Pages[0].Width) * 72.0 / 96.0);
    }

    public void SetPagesPerRow(int n) { PagesPerRow = Math.Clamp(n, 1, 4); InvalidateMeasure(); InvalidateVisual(); }

    public bool FitPage(double viewportW, double viewportH)
    {
        PagesPerRow = 1;
        if (_s == null || _s.Layout.Pages.Count == 0) return false;
        var p = _s.Layout.Pages[0];
        SetZoom(Math.Min((viewportW - 2 * SideDip) / p.Width, (viewportH - 2 * GapDip) / p.Height) * 72.0 / 96.0);
        return true;
    }

    public void FitWidth(double viewportW)
    {
        PagesPerRow = 1;
        if (_s == null || _s.Layout.Pages.Count == 0) return;
        SetZoom((viewportW - 2 * SideDip) / _s.Layout.Pages[0].Width * 72.0 / 96.0);
    }

    // ---- stampa/PDF: stessa pipeline ----
    public LayoutResult? Layout => _s?.Layout;
}
