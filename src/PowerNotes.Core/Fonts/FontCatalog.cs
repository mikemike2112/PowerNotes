using Microsoft.Win32;

namespace PowerNotes.Core.Fonts;

/// <summary>Trova i file font installati (registro Windows) e misura il testo con le metriche reali.</summary>
public static class FontCatalog
{
    public record Entry(string File, int Index);

    static readonly Dictionary<string, Entry> Map = Build();   // chiave: "famiglia|b|i"
    static readonly Dictionary<string, FontFile?> Cache = new();
    const string Fallback = "Arial";

    static Dictionary<string, Entry> Build()
    {
        var m = new Dictionary<string, Entry>(StringComparer.OrdinalIgnoreCase);
        if (!OperatingSystem.IsWindows()) return m;
        var dir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        foreach (var (hive, key) in new[] { (Registry.LocalMachine, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"), (Registry.CurrentUser, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") })
        {
            using var k = hive.OpenSubKey(key);
            if (k == null) continue;
            foreach (var name in k.GetValueNames())
            {
                var file = k.GetValue(name) as string;
                if (file == null || !(file.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) || file.EndsWith(".otf", StringComparison.OrdinalIgnoreCase) || file.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))) continue;
                if (!Path.IsPathRooted(file)) file = Path.Combine(dir, file);
                var title = System.Text.RegularExpressions.Regex.Replace(name, @"\s*\((TrueType|OpenType)\)$", "");
                var parts = title.Split(" & ");
                for (int i = 0; i < parts.Length; i++)
                {
                    var p = parts[i].Trim();
                    bool b = false, it = false;
                    foreach (var suf in new[] { " Bold Italic", " Bold Oblique", " Bold", " Italic", " Oblique", " Regular" })
                        if (p.EndsWith(suf, StringComparison.OrdinalIgnoreCase))
                        {
                            b = suf.Contains("Bold", StringComparison.OrdinalIgnoreCase);
                            it = suf.Contains("Italic", StringComparison.OrdinalIgnoreCase) || suf.Contains("Oblique", StringComparison.OrdinalIgnoreCase);
                            p = p[..^suf.Length];
                            break;
                        }
                    m.TryAdd(Key(p, b, it), new Entry(file, i));
                }
            }
        }
        return m;
    }

    static string Key(string family, bool b, bool i) => $"{family}|{(b ? 1 : 0)}|{(i ? 1 : 0)}";

    public static bool Has(string family) => Map.ContainsKey(Key(family, false, false));

    static readonly HashSet<string> CloudScanned = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Font "cloud" di Office scaricati nella cache utente (es. Aptos): cartella per famiglia, stile dai metadati OS/2.</summary>
    static void ScanCloud(string family)
    {
        if (!OperatingSystem.IsWindows() || !CloudScanned.Add(family) || family.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return;
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "FontCache", "4", "CloudFonts", family);
        if (!Directory.Exists(dir)) return;
        foreach (var f in Directory.EnumerateFiles(dir, "*.ttf"))
        {
            try
            {
                var ff = FontFile.Load(f);
                if (ff.Weight is not (400 or 700)) continue;
                Map.TryAdd(Key(family, ff.Weight == 700, ff.Italic), new Entry(f, 0));
            }
            catch (Exception e) when (e is IOException or KeyNotFoundException or ArgumentException or UnauthorizedAccessException) { }
        }
    }

    /// <summary>File del font che verrà usato per (famiglia, stile): con ripiego su stile base e poi su Arial.</summary>
    public static Entry? Resolve(string family, bool bold, bool italic)
    {
        lock (Cache)
        {
            if (!Map.ContainsKey(Key(family, false, false))) ScanCloud(family);
            foreach (var cand in new[] { (family, bold, italic), (family, bold, false), (family, false, italic), (family, false, false), (Fallback, bold, italic), (Fallback, false, false) })
                if (Map.TryGetValue(Key(cand.Item1, cand.Item2, cand.Item3), out var e)) return e;
            return null;
        }
    }

    public static FontFile Get(string family, bool bold, bool italic)
    {
        var ck = Key(family, bold, italic);
        lock (Cache)
        {
            if (Cache.TryGetValue(ck, out var f) && f != null) return f;
            var e = Resolve(family, bold, italic) ?? throw new InvalidOperationException($"Nessun font disponibile per '{family}' (né il predefinito {Fallback}).");
            f = FontFile.Load(e.File, e.Index); Cache[ck] = f; return f;
        }
    }

    /// <summary>Larghezza in punti. Niente kerning né legature (ponytail: ±1% rispetto a Word).</summary>
    public static double Measure(string text, string family, double sizePt, bool bold, bool italic)
    {
        var f = Get(family, bold, italic);
        double w = 0;
        for (int i = 0; i < text.Length; i++)
        {
            int cp = text[i];
            if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length) cp = char.ConvertToUtf32(text[i], text[++i]);
            w += f.Advance(cp);
        }
        return w * sizePt;
    }

    public static double LineHeight(string family, double sizePt, bool bold, bool italic) => Get(family, bold, italic).LineHeightEm * sizePt;
    public static double AscentOf(string family, double sizePt, bool bold, bool italic) { var f = Get(family, bold, italic); return (f.Ascent + f.LineGap) * sizePt; }
}
