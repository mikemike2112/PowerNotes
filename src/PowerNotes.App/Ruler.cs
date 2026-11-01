using System.Windows;
using System.Windows.Media;

namespace PowerNotes.App;

/// <summary>Righello in centimetri: zero sul margine, area di testo chiara, margini scuri.</summary>
public sealed class Ruler : FrameworkElement
{
    public bool Vertical { get; set; }
    double _origin, _scale = 1.3333, _start, _len, _m0, _m1;

    /// <param name="pageStart">posizione DIP del bordo pagina nella coordinata del righello</param>
    public void Set(double pageStart, double scale, double lenPt, double marginStartPt, double marginEndPt)
    {
        _start = pageStart; _scale = scale; _len = lenPt; _m0 = marginStartPt; _m1 = marginEndPt; InvalidateVisual();
    }

    protected override void OnRender(DrawingContext dc)
    {
        var bg = (Brush)FindResource("BgBrush"); var panel = (Brush)FindResource("PanelBrush"); var alt = (Brush)FindResource("PanelAltBrush");
        var text = (Brush)FindResource("MutedBrush");
        double total = Vertical ? ActualHeight : ActualWidth, thick = Vertical ? ActualWidth : ActualHeight;
        Rect R(double a, double b) => Vertical ? new Rect(0, a, thick, Math.Max(0, b - a)) : new Rect(a, 0, Math.Max(0, b - a), thick);
        dc.DrawRectangle(bg, null, R(0, total));
        if (_len <= 0) return;
        double pa = _start, pb = _start + _len * _scale;
        double ta = _start + _m0 * _scale, tb = pb - _m1 * _scale;
        dc.DrawRectangle(alt, null, R(Math.Max(0, pa), Math.Min(total, pb)));
        dc.DrawRectangle(panel, null, R(Math.Max(0, ta), Math.Min(total, tb)));
        var pen = new Pen(text, 0.8); double ppd = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        double cm = 28.3465 * _scale;
        for (int k = (int)Math.Floor((pa - ta) / (cm / 2)); k <= (int)Math.Ceiling((pb - ta) / (cm / 2)); k++)
        {
            double pos = ta + k * cm / 2;
            if (pos < 0 || pos > total || pos < pa || pos > pb) continue;
            bool whole = k % 2 == 0;
            double len = whole ? thick * 0.45 : thick * 0.25, mid = thick / 2;
            if (Vertical) dc.DrawLine(pen, new Point(mid - len / 2, pos), new Point(mid + len / 2, pos)); else dc.DrawLine(pen, new Point(pos, mid - len / 2), new Point(pos, mid + len / 2));
            if (whole && k != 0)
            {
                var ft = new FormattedText(Math.Abs(k / 2).ToString(), System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Segoe UI"), 10, text, ppd);
                if (Vertical) dc.DrawText(ft, new Point((thick - ft.Width) / 2, pos + 2)); else dc.DrawText(ft, new Point(pos + 2, (thick - ft.Height) / 2));
            }
        }
    }
}
