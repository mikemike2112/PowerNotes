using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.IO;

namespace PowerNotes.App;

/// <summary>Cattura della sola finestra con PrintWindow (nessun input, nessun focus).</summary>
static class DialogShot
{
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("gdi32.dll")] static extern IntPtr CreateCompatibleDC(IntPtr hdc);
    [DllImport("gdi32.dll")] static extern IntPtr CreateDIBSection(IntPtr hdc, ref BITMAPINFO bi, uint usage, out IntPtr bits, IntPtr sec, uint off);
    [DllImport("gdi32.dll")] static extern IntPtr SelectObject(IntPtr hdc, IntPtr o);
    [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr o);
    [DllImport("gdi32.dll")] static extern bool DeleteDC(IntPtr dc);
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] struct BITMAPINFO { public int Size, W, H; public short Planes, Bpp; public int Comp, ImgSize, X, Y, Used, Imp; }

    public static void Capture(Window win, string path)
    {
        var h = new WindowInteropHelper(win).Handle;
        GetWindowRect(h, out var r); int w = r.R - r.L, ht = r.B - r.T;
        var bi = new BITMAPINFO { Size = 40, W = w, H = -ht, Planes = 1, Bpp = 32 };
        var dc = CreateCompatibleDC(IntPtr.Zero);
        var bmp = CreateDIBSection(dc, ref bi, 0, out var bits, IntPtr.Zero, 0);
        var old = SelectObject(dc, bmp);
        PrintWindow(h, dc, 2);
        var buf = new byte[w * ht * 4]; Marshal.Copy(bits, buf, 0, buf.Length);
        SelectObject(dc, old); DeleteObject(bmp); DeleteDC(dc);
        var bs = BitmapSource.Create(w, ht, 96, 96, System.Windows.Media.PixelFormats.Bgra32, null, buf, w * 4);
        var enc = new PngBitmapEncoder(); enc.Frames.Add(BitmapFrame.Create(bs));
        using var fs = File.Create(path); enc.Save(fs);
    }
}
