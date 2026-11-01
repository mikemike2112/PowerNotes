using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerNotes.Core.Layout;

namespace PowerNotes.Core.Model;

public sealed class DocumentLoadException(string message, Exception? inner = null) : Exception(message, inner);

/// <summary>Documento aperto: il modello è il package OOXML originale, non una copia. Salvataggio atomico, undo/redo a istantanee XML.</summary>
public sealed class WordDoc : IDisposable
{
    public WordprocessingDocument Pkg { get; private set; }
    public string? Path { get; private set; }
    public Guid Id { get; } = Guid.NewGuid();
    public bool Dirty { get; private set; }
    public int Version { get; private set; }     // incrementa a ogni modifica: invalida il layout
    readonly MemoryStream _ms;

    WordDoc(WordprocessingDocument pkg, MemoryStream ms, string? path) { Pkg = pkg; _ms = ms; Path = path; }

    public Body Body => Pkg.MainDocumentPart!.Document!.Body!;

    /// <summary>Apre un file .docx. File malformato → DocumentLoadException con messaggio chiaro. Nessuna macro eseguita.</summary>
    public static WordDoc Open(string path)
    {
        byte[] bytes;
        try { bytes = File.ReadAllBytes(path); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { throw new DocumentLoadException($"Impossibile leggere il file: {e.Message}", e); }
        var ms = new MemoryStream();
        ms.Write(bytes);
        ms.Position = 0;
        try
        {
            var pkg = WordprocessingDocument.Open(ms, true, new OpenSettings { AutoSave = false });
            if (pkg.MainDocumentPart?.Document?.Body == null) { pkg.Dispose(); throw new DocumentLoadException("Il file non contiene un documento Word valido (manca il corpo)."); }
            return new WordDoc(pkg, ms, path);
        }
        catch (Exception e) when (e is OpenXmlPackageException or InvalidDataException or FileFormatException or System.Xml.XmlException or ArgumentException or InvalidOperationException)
        {
            ms.Dispose();
            throw new DocumentLoadException("Il file è danneggiato o non è un documento .docx valido: " + e.Message, e);
        }
    }

    public static WordDoc CreateNew()
    {
        var ms = new MemoryStream();
        var pkg = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true);
        var main = pkg.AddMainDocumentPart();
        main.Document = new Document(new Body(
            new Paragraph(),
            new SectionProperties(new PageSize { Width = 11906, Height = 16838 }, new PageMargin { Top = 1417, Right = 1134, Bottom = 1134, Left = 1134, Header = 708, Footer = 708, Gutter = 0 })));
        BuiltIns.EnsureStyle(pkg, "Normal");
        BuiltIns.EnsureStyle(pkg, "TableNormal");
        pkg.Save();
        return new WordDoc(pkg, ms, null);
    }

    public LayoutResult Layout() => LayoutEngine.Layout(Pkg);

    // ---- salvataggio ----
    /// <summary>Scrive su file temporaneo nella stessa cartella poi sostituisce: il sorgente non si corrompe se fallisce.</summary>
    public void Save(string? path = null)
    {
        path ??= Path ?? throw new InvalidOperationException("Nessun percorso di salvataggio.");
        Pkg.Save();
        var tmp = path + "." + Guid.NewGuid().ToString("N")[..8] + ".tmp";
        try
        {
            using (var outMs = new MemoryStream())
            {
                using (var clone = Pkg.Clone(outMs, true)) { clone.Save(); }
                File.WriteAllBytes(tmp, outMs.ToArray());
            }
            // verifica che il file scritto si riapra
            using (var chk = WordprocessingDocument.Open(tmp, false)) { _ = chk.MainDocumentPart!.Document!.Body; }
            if (File.Exists(path)) File.Replace(tmp, path, null); else File.Move(tmp, path);
        }
        finally { if (File.Exists(tmp)) { try { File.Delete(tmp); } catch (IOException) { } } }
        Path = path; Dirty = false;
    }

    public byte[] ToBytes()
    {
        Pkg.Save();
        using var outMs = new MemoryStream();
        using (var clone = Pkg.Clone(outMs, true)) { clone.Save(); }
        return outMs.ToArray();
    }

    // ---- undo / redo ----
    public sealed record Snap(Dictionary<string, string> Xml, int CaretPara, int CaretOff, int AnchorPara, int AnchorOff);
    readonly List<Snap> _undo = new(), _redo = new();
    string? _lastKey;
    public bool CanUndo => _undo.Count > 0;
    public bool CanRedo => _redo.Count > 0;

    IEnumerable<(string key, OpenXmlPart part, Func<OpenXmlPartRootElement?> get, Action<string> set)> Roots()
    {
        var m = Pkg.MainDocumentPart!;
        yield return ("doc", m, () => m.Document, x => m.Document = new Document(x));
        if (m.StyleDefinitionsPart is { } sp) yield return ("styles", sp, () => sp.Styles, x => sp.Styles = new Styles(x));
        if (m.NumberingDefinitionsPart is { } np) yield return ("num", np, () => np.Numbering, x => np.Numbering = new Numbering(x));
        foreach (var h in m.HeaderParts) { var hp = h; yield return ("h:" + m.GetIdOfPart(hp), hp, () => hp.Header, x => hp.Header = new Header(x)); }
        foreach (var f in m.FooterParts) { var fp = f; yield return ("f:" + m.GetIdOfPart(fp), fp, () => fp.Footer, x => fp.Footer = new Footer(x)); }
    }

    Snap Capture(int cp, int co, int ap, int ao)
    {
        var d = new Dictionary<string, string>();
        foreach (var (k, _, get, _) in Roots()) if (get() is { } r) d[k] = r.OuterXml;
        return new Snap(d, cp, co, ap, ao);
    }

    /// <summary>Da chiamare PRIMA di ogni modifica. coalesceKey uguale alla precedente = stessa voce di undo (es. digitazione).</summary>
    public void BeginEdit(int caretPara, int caretOff, int anchorPara, int anchorOff, string? coalesceKey = null)
    {
        if (coalesceKey == null || coalesceKey != _lastKey || _undo.Count == 0)
        {
            _undo.Add(Capture(caretPara, caretOff, anchorPara, anchorOff));
            if (_undo.Count > 200) _undo.RemoveAt(0);
        }
        _lastKey = coalesceKey;
        _redo.Clear();
        Dirty = true; Version++;
    }

    public void Touched() { Dirty = true; Version++; }

    Snap Restore(Snap s, Snap current)
    {
        foreach (var (k, _, _, set) in Roots().ToList()) if (s.Xml.TryGetValue(k, out var xml)) set(xml);
        Version++; Dirty = true; _lastKey = null;
        return current;
    }

    /// <summary>Annulla; restituisce (paragrafo, offset, ancora) del caret salvato.</summary>
    public Snap? Undo(int cp, int co, int ap, int ao)
    {
        if (_undo.Count == 0) return null;
        var s = _undo[^1]; _undo.RemoveAt(_undo.Count - 1);
        _redo.Add(Capture(cp, co, ap, ao));
        Restore(s, s);
        return s;
    }

    public Snap? Redo(int cp, int co, int ap, int ao)
    {
        if (_redo.Count == 0) return null;
        var s = _redo[^1]; _redo.RemoveAt(_redo.Count - 1);
        _undo.Add(Capture(cp, co, ap, ao));
        Restore(s, s);
        return s;
    }

    /// <summary>Tutti i paragrafi del corpo in ordine di documento, celle incluse.</summary>
    public List<Paragraph> Paragraphs() => Body.Descendants<Paragraph>().Where(p => !p.Ancestors<Run>().Any()).ToList();

    public void Dispose() { Pkg.Dispose(); _ms.Dispose(); }
}
