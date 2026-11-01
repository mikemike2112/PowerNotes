using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PowerNotes.Core.Fonts;
using PowerNotes.Core.Layout;

namespace PowerNotes.App;

/// <summary>Disegna una pagina del layout (coordinate in punti) su un DrawingContext. Usato da schermo e stampa: stessa pipeline.</summary>
public static class PageRenderer
{
    static readonly Dictionary<string, GlyphTypeface?> Faces = new();
    static readonly Dictionary<string, Brush> Brushes = new();
    static readonly Dictionary<ImagePrim, BitmapSource?> Images = new(ReferenceEqualityComparer.Instance);

    static Brush Brush(string hex)
    {
        if (Brushes.TryGetValue(hex, out var b)) return b;
        var c = (Color)ColorConverter.ConvertFromString("#" + hex);
        b = new SolidColorBrush(c); b.Freeze();
        return Brushes[hex] = b;
    }

    static GlyphTypeface? Face(string family, bool bold, bool italic)
    {
        var key = $"{family}|{bold}|{italic}";
        if (Faces.TryGetValue(key, out var g)) return g;
        try
        {
            var e = FontCatalog.Resolve(family, bold, italic);
            g = e == null || e.File.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase) ? null : new GlyphTypeface(new Uri(e.File));
        }
        catch (Exception ex) when (ex is IOException or ArgumentException or NotSupportedException or FileFormatException) { g = null; }
        return Faces[key] = g;
    }

    public static void Draw(DrawingContext dc, Page pg, double pixelsPerDip, bool white = true)
    {
        if (white) dc.DrawRectangle(System.Windows.Media.Brushes.White, null, new Rect(0, 0, pg.Width, pg.Height));
        foreach (var p in pg.Prims)
            switch (p)
            {
                case RectPrim r: dc.DrawRectangle(Brush(r.Fill), null, new Rect(r.X, r.Y, r.W, r.H)); break;
                case LinePrim l:
                    var pen = new Pen(Brush(l.Color), l.Width); pen.Freeze();
                    dc.DrawLine(pen, new Point(l.X1, l.Y1), new Point(l.X2, l.Y2)); break;
                case TextPrim t: DrawText(dc, t, pixelsPerDip); break;
                case ImagePrim im: DrawImage(dc, im); break;
            }
    }

    static void DrawText(DrawingContext dc, TextPrim t, double ppd)
    {
        var f = t.Fmt; double size = f.EffSize;
        var brush = Brush(f.Color ?? "000000");
        var face = Face(f.Font, f.Bold, f.Italic);
        if (face != null)
        {
            var map = face.CharacterToGlyphMap;
            var idx = new List<ushort>(t.Text.Length); var adv = new List<double>(t.Text.Length);
            foreach (var ch in t.Text)
            {
                map.TryGetValue(ch, out var gi);
                idx.Add(gi); adv.Add(face.AdvanceWidths.TryGetValue(gi, out var a) ? a * size : 0);
            }
            bool synB = f.Bold && !FontCatalog.HasExact(f.Font, true, f.Italic), synI = f.Italic && !FontCatalog.HasExact(f.Font, f.Bold, true);
            if (synI) dc.PushTransform(new MatrixTransform(1, 0, -0.2, 1, 0.2 * t.Baseline, 0));   // corsivo sintetizzato
            dc.DrawGlyphRun(brush, new GlyphRun(face, 0, false, size, (float)ppd, idx, new Point(t.X, t.Baseline), adv, null, null, null, null, null, null));
            if (synB) dc.DrawGlyphRun(brush, new GlyphRun(face, 0, false, size, (float)ppd, idx, new Point(t.X + size * 0.03, t.Baseline), adv, null, null, null, null, null, null));   // grassetto sintetizzato
            if (synI) dc.Pop();
            return;
        }
        // ripiego (es. collezioni .ttc): FormattedText con il nome di famiglia
        var tf = new Typeface(new FontFamily(f.Font), f.Italic ? FontStyles.Italic : FontStyles.Normal, f.Bold ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal);
        var ft = new FormattedText(t.Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, tf, size, brush, ppd);
        dc.DrawText(ft, new Point(t.X, t.Baseline - ft.Baseline));
    }

    static void DrawImage(DrawingContext dc, ImagePrim im)
    {
        if (!Images.TryGetValue(im, out var bmp))
        {
            try
            {
                if (im.Part != null)
                {
                    using var s = im.Part.GetStream(); var ms = new MemoryStream(); s.CopyTo(ms); ms.Position = 0;
                    var b = new BitmapImage(); b.BeginInit(); b.CacheOption = BitmapCacheOption.OnLoad; b.StreamSource = ms; b.EndInit(); b.Freeze();
                    bmp = b;
                }
            }
            catch (Exception ex) when (ex is IOException or NotSupportedException or InvalidOperationException or FileFormatException) { bmp = null; }
            Images[im] = bmp;
        }
        if (bmp != null) dc.DrawImage(bmp, new Rect(im.X, im.Y, im.W, im.H));
        else dc.DrawRectangle(null, new Pen(System.Windows.Media.Brushes.Gray, 0.75), new Rect(im.X, im.Y, im.W, im.H));
    }
}
