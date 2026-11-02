using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PowerNotes.Core.Fonts;
using PowerNotes.Core.Layout;

namespace PowerNotes.Core.Pdf;

/// <summary>Risolve i font di PDFsharp con lo stesso catalogo usato dal layout (stessi file, stesse metriche).</summary>
sealed class CatalogFontResolver : IFontResolver
{
    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var e = FontCatalog.Resolve(familyName, isBold, isItalic);
        if (e == null) return null;
        bool simB = isBold && !FontCatalog.HasExact(familyName, true, isItalic), simI = isItalic && !FontCatalog.HasExact(familyName, isBold, true);
        // le collezioni .ttc si incorporano estraendo il singolo font (nome faccia = "file#indice")
        return new FontResolverInfo(e.File.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase) ? e.File + "#" + e.Index : e.File, simB, simI);
    }

    public byte[]? GetFont(string faceName)
    {
        int h = faceName.LastIndexOf('#');
        return h > 0 ? PowerNotes.Core.Fonts.FontFile.ExtractFromCollection(faceName[..h], int.Parse(faceName[(h + 1)..])) : File.ReadAllBytes(faceName);
    }
}

/// <summary>Esporta il risultato del layout in PDF: stessa pipeline dello schermo, testo vettoriale con font incorporati.</summary>
public static class PdfExporter
{
    static bool _init;

    static void Init()
    {
        if (_init) return;
        GlobalFontSettings.FontResolver = new CatalogFontResolver();
        _init = true;
    }

    static XColor Col(string? hex)
    {
        if (hex is { Length: 6 } && int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out var v))
            return XColor.FromArgb((v >> 16) & 255, (v >> 8) & 255, v & 255);
        return XColors.Black;
    }

    public static byte[] Export(LayoutResult layout, string title = "")
    {
        Init();
        using var pdf = new PdfDocument();
        pdf.Info.Title = title;
        pdf.Info.Creator = "PowerNotes";
        foreach (var pg in layout.Pages)
        {
            var page = pdf.AddPage();
            page.Width = XUnit.FromPoint(pg.Width); page.Height = XUnit.FromPoint(pg.Height);
            using var g = XGraphics.FromPdfPage(page);
            foreach (var prim in pg.Prims)
                switch (prim)
                {
                    case RectPrim r: g.DrawRectangle(new XSolidBrush(Col(r.Fill)), r.X, r.Y, r.W, r.H); break;
                    case LinePrim l: g.DrawLine(new XPen(Col(l.Color), l.Width), l.X1, l.Y1, l.X2, l.Y2); break;
                    case TextPrim t:
                        {
                            var style = (t.Fmt.Bold ? XFontStyleEx.Bold : XFontStyleEx.Regular) | (t.Fmt.Italic ? XFontStyleEx.Italic : XFontStyleEx.Regular);
                            var font = new XFont(t.Fmt.Font, t.Fmt.EffSize, style, new XPdfFontOptions(PdfFontEncoding.Unicode));
                            g.DrawString(t.Text, font, new XSolidBrush(Col(t.Fmt.Color)), new XPoint(t.X, t.Baseline), XStringFormats.BaseLineLeft);
                            break;
                        }
                    case ShapePrim sh:
                        {
                            var brush = sh.Fill != null ? new XSolidBrush(Col(sh.Fill)) : null;
                            var pen = sh.Line != null ? new XPen(Col(sh.Line), sh.LineW) : null;
                            var poly = Shapes.Polygon(sh.Geom, sh.W, sh.H);
                            if (poly != null)
                            {
                                var pts = poly.Select(q => new XPoint(sh.X + q.X, sh.Y + q.Y)).ToArray();
                                if (brush != null && pen != null) g.DrawPolygon(pen, brush, pts, XFillMode.Winding); else if (brush != null) g.DrawPolygon(brush, pts, XFillMode.Winding); else if (pen != null) g.DrawPolygon(pen, pts);
                            }
                            else if (sh.Geom == "ellipse") { if (brush != null && pen != null) g.DrawEllipse(pen, brush, sh.X, sh.Y, sh.W, sh.H); else if (brush != null) g.DrawEllipse(brush, sh.X, sh.Y, sh.W, sh.H); else if (pen != null) g.DrawEllipse(pen, sh.X, sh.Y, sh.W, sh.H); }
                            else if (sh.Geom == "roundRect") { double r = Shapes.Radius(sh.W, sh.H) * 2; if (brush != null && pen != null) g.DrawRoundedRectangle(pen, brush, sh.X, sh.Y, sh.W, sh.H, r, r); else if (brush != null) g.DrawRoundedRectangle(brush, sh.X, sh.Y, sh.W, sh.H, r, r); else if (pen != null) g.DrawRoundedRectangle(pen, sh.X, sh.Y, sh.W, sh.H, r, r); }
                            else { if (brush != null) g.DrawRectangle(brush, sh.X, sh.Y, sh.W, sh.H); if (pen != null) g.DrawRectangle(pen, sh.X, sh.Y, sh.W, sh.H); }
                            break;
                        }
                    case ImagePrim im when im.Part != null:
                        try
                        {
                            var bytes = new MemoryStream();
                            using (var s = im.Part.GetStream()) s.CopyTo(bytes);
                            bytes.Position = 0;
                            using var img = XImage.FromStream(bytes);
                            g.DrawImage(img, im.X, im.Y, im.W, im.H);
                        }
                        catch (Exception e) when (e is IOException or InvalidOperationException or ArgumentException or NotImplementedException or NotSupportedException)
                        {
                            g.DrawRectangle(XPens.Gray, im.X, im.Y, im.W, im.H);   // immagine illeggibile: segnaposto
                        }
                        break;
                }
        }
        using var ms = new MemoryStream();
        pdf.Save(ms);
        return ms.ToArray();
    }
}
