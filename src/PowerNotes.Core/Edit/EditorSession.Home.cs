using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Model;

namespace PowerNotes.Core.Edit;

public enum CaseMode { Sentence, Lower, Upper, Title, Toggle }
public enum BorderEdge { None, Box, Top, Bottom, Left, Right }

/// <summary>Scheda Home: copia formato, maiuscole/minuscole, ordina paragrafi, sfondo e bordi di paragrafo.</summary>
public sealed partial class EditorSession
{
    // ---------- copia formato ----------
    /// <summary>Formattazione diretta del carattere al cursore (senza lo stile «Collegamento»), da applicare altrove con <see cref="ApplyFormat"/>.</summary>
    public RunProperties CaptureFormat()
    {
        var leaves = ParaText.Leaves(Caret.P);
        var src = Pending ?? leaves.LastOrDefault(l => l.Start < Caret.Off)?.Run.RunProperties ?? leaves.FirstOrDefault()?.Run.RunProperties;
        var rp = src != null ? (RunProperties)src.CloneNode(true) : new RunProperties();
        if (rp.RunStyle?.Val?.Value is { } sid && Styles.Find(sid)?.StyleName?.Val?.Value is "Hyperlink") rp.RunStyle = null;
        return rp;
    }

    /// <summary>Sostituisce la formattazione diretta del testo selezionato (o del testo digitato da qui) con quella catturata.</summary>
    public void ApplyFormat(RunProperties fmt) => ApplyRun(r => { r.RemoveAllChildren(); foreach (var c in fmt.ChildElements) r.Append(c.CloneNode(true)); });

    // ---------- maiuscole / minuscole ----------
    static string CaseTransform(string full, CaseMode mode)
    {
        var o = new char[full.Length];
        bool start = true;
        for (int i = 0; i < full.Length; i++)
        {
            char c = full[i], prev = i > 0 ? full[i - 1] : ' ';
            switch (mode)
            {
                case CaseMode.Lower: o[i] = char.ToLowerInvariant(c); break;
                case CaseMode.Upper: o[i] = char.ToUpperInvariant(c); break;
                case CaseMode.Toggle: o[i] = char.IsUpper(c) ? char.ToLowerInvariant(c) : char.ToUpperInvariant(c); break;
                case CaseMode.Title: o[i] = !char.IsLetterOrDigit(prev) && prev != '\'' && prev != '’' ? char.ToUpperInvariant(c) : char.ToLowerInvariant(c); break;
                default:
                    if (char.IsLetter(c)) { o[i] = start ? char.ToUpperInvariant(c) : char.ToLowerInvariant(c); start = false; }
                    else { o[i] = c; if (c is '.' or '!' or '?') start = true; else if (char.IsDigit(c)) start = false; }
                    break;
            }
        }
        return new string(o);
    }

    /// <summary>Cambia maiuscole/minuscole del testo selezionato (senza selezione: la parola sotto il cursore). Ogni carattere resta della stessa lunghezza, la formattazione non cambia.</summary>
    public void ChangeCase(CaseMode mode)
    {
        var ranges = new List<(Paragraph P, int A, int B)>();
        if (HasSelection)
        {
            var (s, e) = Range;
            foreach (var p in SelectedParagraphs()) ranges.Add((p, p == s.P ? s.Off : 0, p == e.P ? e.Off : ParaText.Length(p)));
        }
        else
        {
            var t = ParaText.Text(Caret.P); int i = Math.Min(Caret.Off, t.Length), a = i, b = i;
            while (a > 0 && char.IsLetterOrDigit(t[a - 1])) a--;
            while (b < t.Length && char.IsLetterOrDigit(t[b])) b++;
            if (b > a) ranges.Add((Caret.P, a, b));
        }
        ranges.RemoveAll(r => r.B <= r.A);
        if (ranges.Count == 0) return;
        Begin();
        foreach (var (p, a, b) in ranges)
        {
            XmlOps.EnsureBoundary(p, b); XmlOps.EnsureBoundary(p, a);
            var changed = CaseTransform(ParaText.Text(p), mode);
            foreach (var l in ParaText.Leaves(p).Where(l => l.Kind == LeafKind.Text && l.Field == null && l.El is Text && l.Start >= a && l.End <= b))
            {
                var t = (Text)l.El;
                t.Text = changed.Substring(l.Start, l.Length); t.Space = SpaceProcessingModeValues.Preserve;
            }
        }
        Changed();
    }

    // ---------- ordina paragrafi ----------
    /// <summary>Ordina alfabeticamente i paragrafi selezionati (consecutivi, dello stesso blocco).</summary>
    public void SortParagraphs(bool descending)
    {
        var ps = SelectedParagraphs().ToList();
        if (ps.Count < 2) throw new InvalidOperationException("Seleziona almeno due paragrafi da ordinare.");
        for (int i = 0; i < ps.Count - 1; i++)
            if (ps[i].NextSibling() != ps[i + 1]) throw new InvalidOperationException("Ordina funziona su paragrafi consecutivi dello stesso blocco (non attraverso tabelle).");
        if (ps.Any(p => p.ParagraphProperties?.SectionProperties != null)) throw new InvalidOperationException("Ordina non è disponibile attraverso un'interruzione di sezione.");
        var keys = ps.ToDictionary(p => p, p => ParaText.Text(p).Trim());
        var cmp = Comparer<Paragraph>.Create((x, y) => string.Compare(keys[x], keys[y], StringComparison.InvariantCultureIgnoreCase));
        var sorted = (descending ? ps.OrderByDescending(p => p, cmp) : ps.OrderBy(p => p, cmp)).ToList();
        if (sorted.SequenceEqual(ps)) return;
        Begin();
        var parent = ps[0].Parent!; var after = ps[^1].NextSibling();
        foreach (var p in ps) p.Remove();
        foreach (var p in sorted) { if (after != null) after.InsertBeforeSelf(p); else parent.Append(p); }
        Changed();
    }

    // ---------- sfondo e bordi di paragrafo ----------
    /// <summary>Sfondo dei paragrafi selezionati (RRGGBB) o null per toglierlo.</summary>
    public void SetParaShading(string? hex)
    {
        if (hex != null && !Regex.IsMatch(hex, "^[0-9A-Fa-f]{6}$")) throw new ArgumentException("Colore non valido.", nameof(hex));
        ApplyPara(pp => pp.Shading = hex == null ? null : new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = hex.ToUpperInvariant() });
    }

    static T Edge<T>(int space) where T : BorderType, new() => new() { Val = BorderValues.Single, Size = 4, Space = (uint)space, Color = "auto" };

    /// <summary>Bordi dei paragrafi selezionati: nessuno, contorno, oppure un lato (che si attiva o disattiva).</summary>
    public void SetParaBorders(BorderEdge edge)
    {
        var ps = SelectedParagraphs().ToList();
        bool Has(Paragraph p, BorderEdge e) => p.ParagraphProperties?.ParagraphBorders is { } b && (e switch { BorderEdge.Top => b.TopBorder, BorderEdge.Bottom => b.BottomBorder, BorderEdge.Left => b.LeftBorder, _ => (BorderType?)b.RightBorder }) != null;
        bool remove = edge is not (BorderEdge.None or BorderEdge.Box) && ps.All(p => Has(p, edge));
        ApplyPara(pp =>
        {
            if (edge == BorderEdge.None) { pp.ParagraphBorders = null; return; }
            var b = pp.ParagraphBorders ?? (pp.ParagraphBorders = new ParagraphBorders());
            void Put(BorderEdge e, bool on)
            {
                switch (e)
                {
                    case BorderEdge.Top: b.TopBorder = on ? Edge<TopBorder>(1) : null; break;
                    case BorderEdge.Left: b.LeftBorder = on ? Edge<LeftBorder>(4) : null; break;
                    case BorderEdge.Bottom: b.BottomBorder = on ? Edge<BottomBorder>(1) : null; break;
                    case BorderEdge.Right: b.RightBorder = on ? Edge<RightBorder>(4) : null; break;
                }
            }
            if (edge == BorderEdge.Box) foreach (var e in new[] { BorderEdge.Top, BorderEdge.Left, BorderEdge.Bottom, BorderEdge.Right }) Put(e, true);
            else Put(edge, !remove);
            if (!b.HasChildren) pp.ParagraphBorders = null;
        });
    }
}
