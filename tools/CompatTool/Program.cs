// Strumento per il confronto visivo (COMPAT.md): PDF dal clone, rasterizzazione con Windows.Data.Pdf, affiancamento Word | clone.
//   CompatTool clone-pdf <file.docx> <out.pdf>
//   CompatTool compare <word.pdf> <clone.pdf> <out.png> [pagina=1] [larghezza=700]
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PowerNotes.Core.Model;
using PowerNotes.Core.Pdf;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;
using System.Runtime.InteropServices.WindowsRuntime;

static class Program
{
    [STAThread]
    static async Task<int> Main(string[] a)
    {
        try
        {
            if (a.Length >= 3 && a[0] == "clone-pdf")
            {
                using var d = WordDoc.Open(a[1]);
                File.WriteAllBytes(a[2], PdfExporter.Export(d.Layout(), Path.GetFileNameWithoutExtension(a[1])));
                return 0;
            }
            if (a.Length >= 4 && a[0] == "compare")
            {
                int page = a.Length > 4 ? int.Parse(a[4]) : 1; int width = a.Length > 5 ? int.Parse(a[5]) : 700;
                var (l, ln) = await Raster(a[1], page, width); var (r, rn) = await Raster(a[2], page, width);
                Compose(l, r, a[3]);
                Console.WriteLine($"pagine Word={ln} clone={rn}");
                return 0;
            }
            Console.Error.WriteLine("uso: clone-pdf <docx> <pdf> | compare <word.pdf> <clone.pdf> <out.png> [pagina] [larghezza]");
            return 2;
        }
        catch (Exception e) { Console.Error.WriteLine(e.Message); return 1; }
    }

    static async Task<(BitmapSource, uint)> Raster(string pdf, int page, int width)
    {
        var file = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(pdf));
        var doc = await PdfDocument.LoadFromFileAsync(file);
        if (page > doc.PageCount) throw new InvalidOperationException($"{pdf}: solo {doc.PageCount} pagine");
        using var p = doc.GetPage((uint)(page - 1));
        using var ms = new InMemoryRandomAccessStream();
        await p.RenderToStreamAsync(ms, new PdfPageRenderOptions { DestinationWidth = (uint)width });
        var bmp = new BitmapImage(); bmp.BeginInit(); bmp.CacheOption = BitmapCacheOption.OnLoad; bmp.StreamSource = ms.AsStreamForRead(); bmp.EndInit(); bmp.Freeze();
        return (bmp, doc.PageCount);
    }

    static void Compose(BitmapSource l, BitmapSource r, string path)
    {
        int h = Math.Max(l.PixelHeight, r.PixelHeight), w = l.PixelWidth + r.PixelWidth + 30;
        var dv = new DrawingVisual();
        using (var dc = dv.RenderOpen())
        {
            dc.DrawRectangle(Brushes.DimGray, null, new Rect(0, 0, w, h));
            dc.DrawImage(l, new Rect(0, 0, l.PixelWidth, l.PixelHeight));
            dc.DrawImage(r, new Rect(l.PixelWidth + 30, 0, r.PixelWidth, r.PixelHeight));
        }
        var rtb = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32); rtb.Render(dv);
        var enc = new PngBitmapEncoder(); enc.Frames.Add(BitmapFrame.Create(rtb));
        using var fs = File.Create(path); enc.Save(fs);
    }
}
