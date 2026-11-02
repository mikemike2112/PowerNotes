namespace PowerNotes.Core.Layout;

/// <summary>Forme di base (geometrie predefinite OOXML, regolazioni di default) condivise da schermo, PDF e comando Inserisci forma.</summary>
public static class Shapes
{
    public static readonly (string Prst, string Name)[] Gallery =
    [
        ("rect", "Rettangolo"), ("roundRect", "Rettangolo arrotondato"), ("ellipse", "Ellisse"), ("triangle", "Triangolo"),
        ("rtTriangle", "Triangolo rettangolo"), ("diamond", "Rombo"), ("hexagon", "Esagono"), ("rightArrow", "Freccia a destra"), ("star5", "Stella a cinque punte"),
    ];

    public static bool IsKnown(string prst) => Gallery.Any(g => g.Prst == prst);

    /// <summary>Raggio degli angoli del rettangolo arrotondato (adj di default 16667).</summary>
    public static double Radius(double w, double h) => Math.Min(w, h) * 0.16667;

    /// <summary>Vertici del poligono (coordinate relative alla forma) per le geometrie poligonali; null per rect, roundRect, ellipse.</summary>
    public static (double X, double Y)[]? Polygon(string prst, double w, double h) => prst switch
    {
        "triangle" => [(w / 2, 0), (w, h), (0, h)],
        "rtTriangle" => [(0, 0), (w, h), (0, h)],
        "diamond" => [(w / 2, 0), (w, h / 2), (w / 2, h), (0, h / 2)],
        "hexagon" => [(w * 0.25, 0), (w * 0.75, 0), (w, h / 2), (w * 0.75, h), (w * 0.25, h), (0, h / 2)],
        "rightArrow" => ArrowPts(w, h),
        "star5" => Star(w, h),
        _ => null,
    };

    static (double, double)[] ArrowPts(double w, double h)
    {
        double hl = Math.Min(h, w) * 0.5;
        return [(0, h / 4), (w - hl, h / 4), (w - hl, 0), (w, h / 2), (w - hl, h), (w - hl, h * 3 / 4), (0, h * 3 / 4)];
    }

    static (double, double)[] Star(double w, double h)
    {
        var pts = new (double, double)[10];
        for (int i = 0; i < 10; i++)
        {
            double r = i % 2 == 0 ? 1 : 0.382, a = -Math.PI / 2 + i * Math.PI / 5;
            pts[i] = (w / 2 + Math.Cos(a) * r * w / 2, h / 2 + Math.Sin(a) * r * h / 2 * (i % 2 == 0 ? 1 : 1));
        }
        return pts;
    }
}
