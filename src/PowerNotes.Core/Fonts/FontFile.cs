using System.Buffers.Binary;

namespace PowerNotes.Core.Fonts;

/// <summary>Metriche reali di un font TrueType/OpenType (glifi TrueType o CFF): cmap, hmtx, hhea. Niente kerning.</summary>
public sealed class FontFile
{
    public int UnitsPerEm { get; private set; } = 1000;
    public double Ascent { get; private set; }   // in em
    public double Descent { get; private set; }  // in em, positivo
    public double LineGap { get; private set; }  // in em

    public int Weight { get; private set; } = 400;
    public bool Italic { get; private set; }
    byte[] _d = [];
    int _numHMetrics;
    int _hmtx;
    int _numGlyphs;
    readonly Dictionary<int, ushort> _cmap = new();
    readonly Dictionary<int, double> _adv = new();

    /// <summary>Altezza riga "single" come Word: ascent + descent + lineGap (hhea).</summary>
    public double LineHeightEm => Ascent + Descent + LineGap;

    public static FontFile Load(string path, int index = 0)
    {
        var f = new FontFile { _d = File.ReadAllBytes(path) };
        f.Parse(index);
        return f;
    }

    ushort U16(int o) => BinaryPrimitives.ReadUInt16BigEndian(_d.AsSpan(o));
    short S16(int o) => BinaryPrimitives.ReadInt16BigEndian(_d.AsSpan(o));
    uint U32(int o) => BinaryPrimitives.ReadUInt32BigEndian(_d.AsSpan(o));

    void Parse(int index)
    {
        int off = 0;
        if (U32(0) == 0x74746366) off = (int)U32(12 + 4 * index);   // 'ttcf'
        int n = U16(off + 4);
        var tables = new Dictionary<string, int>();
        for (int i = 0; i < n; i++)
        {
            int e = off + 12 + 16 * i;
            tables[System.Text.Encoding.ASCII.GetString(_d, e, 4)] = (int)U32(e + 8);
        }
        UnitsPerEm = U16(tables["head"] + 18);
        int hhea = tables["hhea"];
        Ascent = (double)S16(hhea + 4) / UnitsPerEm;
        Descent = (double)-S16(hhea + 6) / UnitsPerEm;
        LineGap = (double)S16(hhea + 8) / UnitsPerEm;
        _numHMetrics = U16(hhea + 34);
        _hmtx = tables["hmtx"];
        _numGlyphs = U16(tables["maxp"] + 4);
        if (tables.TryGetValue("OS/2", out var os2) && os2 + 64 <= _d.Length) { Weight = U16(os2 + 4); Italic = (U16(os2 + 62) & 1) != 0; }
        ParseCmap(tables["cmap"]);
    }

    void ParseCmap(int cmap)
    {
        int n = U16(cmap + 2);
        int best = -1, bestScore = -1;
        for (int i = 0; i < n; i++)
        {
            int plat = U16(cmap + 4 + 8 * i), enc = U16(cmap + 6 + 8 * i);
            int score = (plat, enc) switch { (3, 10) => 4, (0, 4) or (0, 6) => 3, (3, 1) => 2, (0, _) => 1, (3, 0) => 1, _ => -1 };
            if (score > bestScore) { bestScore = score; best = cmap + (int)U32(cmap + 8 + 8 * i); }
        }
        if (best < 0) return;
        int fmt = U16(best);
        if (fmt == 4)
        {
            int segX2 = U16(best + 6);
            int endO = best + 14, startO = endO + segX2 + 2, deltaO = startO + segX2, rangeO = deltaO + segX2;
            for (int s = 0; s < segX2 / 2; s++)
            {
                int end = U16(endO + 2 * s), start = U16(startO + 2 * s), delta = S16(deltaO + 2 * s), ro = U16(rangeO + 2 * s);
                if (start == 0xFFFF) continue;
                for (int c = start; c <= end; c++)
                {
                    int g = ro == 0 ? (c + delta) & 0xFFFF : U16(rangeO + 2 * s + ro + 2 * (c - start));
                    if (ro != 0 && g != 0) g = (g + delta) & 0xFFFF;
                    if (g != 0) _cmap[c] = (ushort)g;
                }
            }
        }
        else if (fmt == 12)
        {
            uint groups = U32(best + 12);
            for (uint i = 0; i < groups; i++)
            {
                int g = best + 16 + 12 * (int)i;
                uint s = U32(g), e = U32(g + 4), gid = U32(g + 8);
                for (uint c = s; c <= e && c - s < 70000; c++) _cmap[(int)c] = (ushort)(gid + (c - s));
            }
        }
    }

    public bool HasGlyph(int cp) => _cmap.ContainsKey(cp);

    /// <summary>Avanzamento di un code point in em.</summary>
    public double Advance(int cp)
    {
        if (_adv.TryGetValue(cp, out var a)) return a;
        _cmap.TryGetValue(cp, out var gid);
        int idx = Math.Min((int)gid, _numHMetrics - 1);
        a = (double)U16(_hmtx + 4 * idx) / UnitsPerEm;
        _adv[cp] = a;
        return a;
    }
}
