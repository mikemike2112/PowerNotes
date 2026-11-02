using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Fonts;
using PowerNotes.Core.Model;

namespace PowerNotes.Core.Layout;

// ---- primitive di disegno (punti, origine in alto a sinistra della pagina) ----
public abstract class Prim { }
public sealed class TextPrim : Prim { public double X, Baseline, Width; public string Text = ""; public RunFmt Fmt = new(); }
public sealed class RectPrim : Prim { public double X, Y, W, H; public string Fill = "000000"; }
public sealed class LinePrim : Prim { public double X1, Y1, X2, Y2, Width; public string Color = "000000"; }
public sealed class ImagePrim : Prim { public double X, Y, W, H; public ImagePart? Part; }

/// <summary>Frammento di riga (parola, spazio, tab…) con offset nel paragrafo, usato per caret e hit-test.</summary>
public sealed class SegBox
{
    public int Start, Length;
    public double X, Width;
    public string Text = "";
    public RunFmt Fmt = new();
    public LeafKind Kind;
    public bool Synthetic;   // etichetta di numerazione: non mappata a offset
}

public enum Story { Body, Header, Footer }

public sealed class LineBox
{
    public Story Story;
    public required Paragraph Para;
    public int PageIndex;
    public int Start, End;
    public double Y, Height, Baseline, Left, Right;
    public List<SegBox> Segs = new();
    public bool LastOfPara;

    /// <summary>Ascissa (assoluta sulla pagina) di un offset del paragrafo su questa riga.</summary>
    public double XAt(int offset)
    {
        foreach (var s in Segs)
        {
            if (s.Synthetic) continue;
            if (offset < s.Start || offset > s.Start + s.Length) continue;
            int k = offset - s.Start;
            if (k <= 0) return s.X;
            if (k >= s.Length) return s.X + s.Width;
            if (s.Kind != LeafKind.Text) return s.X;
            if (s.Text.Length == s.Length) return s.X + Metrics.Width(s.Text[..k], s.Fmt);
            return s.X + s.Width * k / s.Length;
        }
        return offset <= Start ? Left : Right;
    }

    /// <summary>Offset del paragrafo più vicino all'ascissa x.</summary>
    public int OffsetAt(double x)
    {
        int best = Start;
        double bestD = double.MaxValue;
        foreach (var s in Segs)
        {
            if (s.Synthetic) continue;
            if (s.Kind != LeafKind.Text || s.Length <= 1 || s.Text.Length != s.Length)
            {
                foreach (var (o, px) in new[] { (s.Start, s.X), (s.Start + s.Length, s.X + s.Width) })
                    if (Math.Abs(px - x) < bestD) { bestD = Math.Abs(px - x); best = o; }
                continue;
            }
            for (int k = 0; k <= s.Length; k++)
            {
                double px = k == 0 ? s.X : k == s.Length ? s.X + s.Width : s.X + Metrics.Width(s.Text[..k], s.Fmt);
                double d = Math.Abs(px - x);
                if (d < bestD) { bestD = d; best = s.Start + k; }
            }
        }
        // a fine riga spezzata l'ultima posizione utile è prima degli spazi finali/a capo
        return Math.Clamp(best, Start, End);
    }
}

public sealed class Page
{
    public int Index, Number;
    public double Width, Height, MarL, MarR, MarT, MarB;
    public List<Prim> Prims = new();
    public List<LineBox> Lines = new();
    /// <summary>Rettangoli delle celle di tabella (per la griglia di visualizzazione).</summary>
    public List<(double X, double Y, double W, double H)> Cells = new();
}

public sealed class LayoutResult
{
    public List<Page> Pages = new();
    readonly Dictionary<Paragraph, List<LineBox>> _byPara = new();

    internal void Register(LineBox l) { if (!_byPara.TryGetValue(l.Para, out var v)) _byPara[l.Para] = v = new(); v.Add(l); }

    public IReadOnlyList<LineBox> LinesOf(Paragraph p) => _byPara.TryGetValue(p, out var v) ? v : [];

    /// <summary>Riga che contiene la posizione. preferPrev: a un confine di riga spezzata, resta sulla riga precedente.</summary>
    public LineBox? Locate(Paragraph p, int offset, bool preferPrev = false)
    {
        var ls = LinesOf(p);
        if (ls.Count == 0) return null;
        LineBox? hit = null;
        for (int i = 0; i < ls.Count; i++)
        {
            var l = ls[i];
            bool last = i == ls.Count - 1;
            if (offset >= l.Start && (offset < l.End || (offset == l.End && (last || preferPrev)))) { hit = l; break; }
            if (offset == l.End && !last && !preferPrev) { continue; }
        }
        return hit ?? ls[^1];
    }

    /// <summary>Hit-test: pagina e punto → (paragrafo, offset).</summary>
    public (Paragraph Para, int Offset, LineBox Line)? HitTest(int pageIndex, double x, double y, bool headerFooter = false)
    {
        if (pageIndex < 0 || pageIndex >= Pages.Count) return null;
        var pg = Pages[pageIndex];
        LineBox? best = null; double bestD = double.MaxValue;
        foreach (var l in pg.Lines)
        {
            if ((l.Story != Story.Body) != headerFooter) continue;
            double d = y < l.Y ? l.Y - y : y > l.Y + l.Height ? y - (l.Y + l.Height) : 0;
            double dx = x < l.Left ? l.Left - x : x > l.Right ? x - l.Right : 0;
            d += dx * 0.01;     // a parità di quota preferisce la cella/colonna giusta
            if (dx > 0 && d < 1e6) d += dx * 0.5;
            if (d < bestD) { bestD = d; best = l; }
        }
        if (best == null) return null;
        return (best.Para, best.OffsetAt(x), best);
    }
}

public static class Metrics
{
    public static double Width(string text, RunFmt f) => FontCatalog.Measure(text, f.Font, f.EffSize, f.Bold, f.Italic);
}
