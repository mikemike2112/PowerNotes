using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace PowerNotes.Core.Privacy;

public enum PrivacyKind { Author, LastAuthor, Company, Manager, CustomProperty, CommentAuthor, RevisionAuthor, Person, PresenceInfo, LocalPath, Thumbnail }
public enum CommentMode { Anonymize, Remove }

/// <param name="Id">identificatore stabile (usato per la selezione)</param>
public sealed record Finding(string Id, PrivacyKind Kind, string Label, string Value, string Part);
public sealed record ScanResult(List<Finding> Findings, List<string> Warnings);
public sealed record CleanResult(byte[] Copy, List<string> Done, ScanResult After);

/// <summary>Scudo privacy per .docx: analizza metadati e autori, crea una COPIA ripulita. Non tocca mai l'input (lavora su byte in memoria). Tutto in locale.</summary>
public static class PrivacyShield
{
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    static readonly XNamespace CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    static readonly XNamespace DC = "http://purl.org/dc/elements/1.1/";
    static readonly XNamespace EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    static readonly XNamespace CUS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    static readonly Regex PathRx = new(@"(?i)(?:file:///?)?[a-z]:[\\/]+users[\\/]+(?<u>[^\\/""<>\s%]+)|\\\\(?<srv>[^\\/""<>\s]+)\\", RegexOptions.Compiled);

    static bool IsBody(string n) => Regex.IsMatch(n, @"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", RegexOptions.IgnoreCase);
    static bool IsComments(string n) => Regex.IsMatch(n, @"^word/comments(Extended|Ids|Extensible)?\.xml$", RegexOptions.IgnoreCase);

    static Dictionary<string, byte[]> ReadAll(byte[] docx, out List<string> order)
    {
        var d = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase); order = new();
        using var z = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);
        foreach (var e in z.Entries.Where(e => !e.FullName.EndsWith('/')))
        {
            using var s = e.Open(); using var ms = new MemoryStream(); s.CopyTo(ms);
            d[e.FullName] = ms.ToArray(); order.Add(e.FullName);
        }
        return d;
    }

    static XDocument? Xml(Dictionary<string, byte[]> p, string name)
    {
        if (!p.TryGetValue(name, out var b)) return null;
        try { using var ms = new MemoryStream(b); return XDocument.Load(ms); } catch (System.Xml.XmlException) { return null; }
    }

    static byte[] Bytes(XDocument d) { using var ms = new MemoryStream(); using (var w = System.Xml.XmlWriter.Create(ms, new System.Xml.XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false), OmitXmlDeclaration = false })) d.Save(w); return ms.ToArray(); }

    // ------------------------------------------------------------ analisi
    public static ScanResult Scan(byte[] docx)
    {
        var p = ReadAll(docx, out _);
        var f = new List<Finding>(); var warn = new List<string>();
        void Add(string id, PrivacyKind k, string label, string value, string part) { if (Regex.IsMatch(value, @"^Autore \d+$") || value == "Corpus") return; /* già anonimo */ if (f.All(x => x.Id != id)) f.Add(new Finding(id, k, label, value, part)); }

        var core = Xml(p, "docProps/core.xml");
        if (core?.Root?.Element(DC + "creator")?.Value is { Length: > 0 } cr) Add("author", PrivacyKind.Author, "Autore", cr, "docProps/core.xml");
        if (core?.Root?.Element(CP + "lastModifiedBy")?.Value is { Length: > 0 } lm) Add("lastauthor", PrivacyKind.LastAuthor, "Ultimo autore", lm, "docProps/core.xml");
        var app = Xml(p, "docProps/app.xml");
        if (app?.Root?.Element(EP + "Company")?.Value is { Length: > 0 } co) Add("company", PrivacyKind.Company, "Società", co, "docProps/app.xml");
        if (app?.Root?.Element(EP + "Manager")?.Value is { Length: > 0 } mg) Add("manager", PrivacyKind.Manager, "Manager", mg, "docProps/app.xml");
        var cust = Xml(p, "docProps/custom.xml");
        foreach (var pr in cust?.Root?.Elements(CUS + "property") ?? [])
            Add("custom:" + pr.Attribute("name")?.Value, PrivacyKind.CustomProperty, "Proprietà personalizzata «" + pr.Attribute("name")?.Value + "»", string.Concat(pr.Elements().Select(e => e.Value)), "docProps/custom.xml");

        foreach (var name in p.Keys.Where(IsComments))
        {
            var x = Xml(p, name); if (x == null) continue;
            foreach (var g in x.Descendants(W + "comment").GroupBy(c => c.Attribute(W + "author")?.Value ?? "").Where(g => g.Key.Length > 0))
                Add("comment:" + g.Key, PrivacyKind.CommentAuthor, $"Autore di commenti ({g.Count()})", g.Key, name);
        }
        foreach (var name in p.Keys.Where(IsBody))
        {
            var x = Xml(p, name); if (x == null) continue;
            foreach (var g in x.Descendants().Where(e => e.Name.Namespace == W && e.Name.LocalName is not "comment" && e.Attribute(W + "author") != null && e.Attribute(W + "id") != null).GroupBy(e => e.Attribute(W + "author")!.Value).Where(g => g.Key.Length > 0))
                Add("rev:" + g.Key, PrivacyKind.RevisionAuthor, $"Autore di revisioni ({g.Count()})", g.Key, name);
            if (x.Descendants(W + "vanish").Any()) { var m = "Testo nascosto (formattazione «nascosto»): il contenuto resta nel file."; if (!warn.Contains(m)) warn.Add(m); }
        }
        var people = Xml(p, "word/people.xml");
        foreach (var per in people?.Descendants(W15 + "person") ?? [])
        {
            if (per.Attribute(W15 + "author")?.Value is { Length: > 0 } a) Add("person:" + a, PrivacyKind.Person, "Persona nell'elenco autori", a, "word/people.xml");
            if (per.Element(W15 + "presenceInfo")?.Attribute(W15 + "userId")?.Value is { Length: > 0 } uid) Add("presence:" + uid, PrivacyKind.PresenceInfo, "Account associato all'autore", uid, "word/people.xml");
        }
        foreach (var (name, bytes) in p.Where(kv => (kv.Key.EndsWith(".xml") || kv.Key.EndsWith(".rels")) && !IsBody(kv.Key) && kv.Key != "[Content_Types].xml"))
        {
            var text = System.Text.Encoding.UTF8.GetString(bytes);
            foreach (Match m in PathRx.Matches(text).Where(m => m.Groups["u"].Value != "utente" && m.Groups["srv"].Value != "server")) Add("path:" + name + ":" + m.Value, PrivacyKind.LocalPath, m.Groups["u"].Success ? "Percorso locale con nome utente" : "Percorso di rete", m.Value, name);
        }
        foreach (var t in p.Keys.Where(k => k.StartsWith("docProps/thumbnail", StringComparison.OrdinalIgnoreCase))) Add("thumb", PrivacyKind.Thumbnail, "Anteprima miniatura del documento", t, t);

        // avvisi su ciò che lo scudo NON sa pulire
        foreach (var (name, b) in p.Where(kv => kv.Key.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase)))
            if (HasExif(b)) { var m = $"L'immagine «{name}» contiene metadati EXIF (posizione, dispositivo, data): non vengono rimossi."; if (!warn.Contains(m)) warn.Add(m); }
        if (p.Keys.Any(k => k.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase))) warn.Add("Il documento contiene oggetti incorporati (word/embeddings): non vengono analizzati né puliti.");
        if (p.Keys.Any(k => k.Contains("vbaProject", StringComparison.OrdinalIgnoreCase) || k.StartsWith("word/activeX", StringComparison.OrdinalIgnoreCase))) warn.Add("Il documento contiene macro o controlli ActiveX: non vengono analizzati né rimossi.");
        warn.Add("Il testo del documento (corpo, intestazioni, note) non viene analizzato: nomi, indirizzi o dati scritti nel contenuto restano.");
        warn.Add("Le revisioni restano nel file (solo l'autore viene anonimizzato): lo scudo non le accetta né le rifiuta.");
        return new ScanResult(f, warn);
    }

    static bool HasExif(byte[] b)
    {
        int n = Math.Min(b.Length, 65536);
        if (n > 4 && b[0] == 0xFF && b[1] == 0xD8) return System.Text.Encoding.Latin1.GetString(b, 0, n).Contains("Exif\0\0");
        if (n > 8 && b[0] == 0x89 && b[1] == (byte)'P') return System.Text.Encoding.Latin1.GetString(b, 0, n).Contains("eXIf");
        return false;
    }

    // ------------------------------------------------------------ pulizia (copia)
    /// <param name="selected">Id dei finding da pulire</param>
    public static CleanResult Clean(byte[] docx, ISet<string> selected, CommentMode comments = CommentMode.Anonymize)
    {
        var scan = Scan(docx);
        var p = ReadAll(docx, out var order);
        var done = new List<string>();
        var sel = scan.Findings.Where(x => selected.Contains(x.Id)).ToList();

        // ---- proprietà
        if (sel.Any(x => x.Id is "author" or "lastauthor") && Xml(p, "docProps/core.xml") is { } core)
        {
            if (selected.Contains("author")) { core.Root!.Element(DC + "creator")!.Value = ""; done.Add("Autore azzerato"); }
            if (selected.Contains("lastauthor")) { core.Root!.Element(CP + "lastModifiedBy")!.Value = ""; done.Add("Ultimo autore azzerato"); }
            p["docProps/core.xml"] = Bytes(core);
        }
        if (sel.Any(x => x.Id is "company" or "manager") && Xml(p, "docProps/app.xml") is { } app)
        {
            if (selected.Contains("company")) { app.Root!.Element(EP + "Company")!.Value = ""; done.Add("Società azzerata"); }
            if (selected.Contains("manager")) { app.Root!.Element(EP + "Manager")!.Value = ""; done.Add("Manager azzerato"); }
            p["docProps/app.xml"] = Bytes(app);
        }
        if (sel.Any(x => x.Kind == PrivacyKind.CustomProperty) && Xml(p, "docProps/custom.xml") is { } cust)
        {
            foreach (var f in sel.Where(x => x.Kind == PrivacyKind.CustomProperty)) { cust.Root!.Elements(CUS + "property").Where(e => "custom:" + e.Attribute("name")?.Value == f.Id).Remove(); done.Add($"Proprietà «{f.Id[7..]}» rimossa"); }
            p["docProps/custom.xml"] = Bytes(cust);
        }

        // ---- autori: pseudonimi coerenti in tutto il file
        var names = sel.Where(x => x.Kind is PrivacyKind.CommentAuthor or PrivacyKind.RevisionAuthor or PrivacyKind.Person).Select(x => x.Value).Distinct().ToList();
        var map = names.Select((n, i) => (n, i)).ToDictionary(t => t.n, t => "Autore " + (t.i + 1));
        var commentAuthors = sel.Where(x => x.Kind == PrivacyKind.CommentAuthor).Select(x => x.Value).ToHashSet();
        var removedCommentIds = new HashSet<string>(); var removedParaIds = new HashSet<string>();
        if (comments == CommentMode.Remove && commentAuthors.Count > 0)
            foreach (var name in p.Keys.Where(k => Regex.IsMatch(k, @"^word/comments\.xml$", RegexOptions.IgnoreCase)).ToList())
            {
                var x = Xml(p, name); if (x == null) continue;
                foreach (var c in x.Descendants(W + "comment").Where(c => commentAuthors.Contains(c.Attribute(W + "author")?.Value ?? "")).ToList())
                {
                    removedCommentIds.Add(c.Attribute(W + "id")!.Value);
                    foreach (var para in c.Descendants(W + "p")) if (para.Attribute(W14 + "paraId")?.Value is { } pid) removedParaIds.Add(pid);
                    c.Remove();
                }
                p[name] = Bytes(x);
            }
        foreach (var name in p.Keys.ToList())
        {
            bool body = IsBody(name), com = IsComments(name), ppl = name.Equals("word/people.xml", StringComparison.OrdinalIgnoreCase);
            if (!body && !com && !ppl) continue;
            var x = Xml(p, name); if (x == null) continue; bool ch = false;
            if (removedCommentIds.Count > 0 && body)
            {
                foreach (var e in x.Descendants().Where(e => e.Name.Namespace == W && e.Name.LocalName is "commentRangeStart" or "commentRangeEnd" && removedCommentIds.Contains(e.Attribute(W + "id")?.Value ?? "")).ToList()) { e.Remove(); ch = true; }
                foreach (var e in x.Descendants(W + "commentReference").Where(e => removedCommentIds.Contains(e.Attribute(W + "id")?.Value ?? "")).ToList()) { (e.Parent?.Name == W + "r" ? e.Parent : e).Remove(); ch = true; }
            }
            if (removedParaIds.Count > 0 && name.EndsWith("commentsExtended.xml", StringComparison.OrdinalIgnoreCase))
                foreach (var e in x.Descendants().Where(e => e.Name.LocalName == "commentEx" && removedParaIds.Contains(e.Attribute(W15 + "paraId")?.Value ?? "")).ToList()) { e.Remove(); ch = true; }
            if (removedParaIds.Count > 0 && name.EndsWith("commentsIds.xml", StringComparison.OrdinalIgnoreCase))
                foreach (var e in x.Descendants().Where(e => e.Name.LocalName == "commentId" && removedParaIds.Contains(e.Attribute(W16cid + "paraId")?.Value ?? "")).ToList()) { e.Remove(); ch = true; }
            foreach (var e in x.Descendants().Where(e => e.Name.Namespace == W && e.Attribute(W + "author") is { } a && map.ContainsKey(a.Value) && (com ? commentAuthors.Contains(a.Value) || true : sel.Any(s => s.Kind == PrivacyKind.RevisionAuthor && s.Value == a.Value))).ToList())
            {
                var at = e.Attribute(W + "author")!;
                bool isComment = e.Name.LocalName == "comment";
                if (isComment ? !commentAuthors.Contains(at.Value) : !sel.Any(s => s.Kind == PrivacyKind.RevisionAuthor && s.Value == at.Value)) continue;
                at.Value = map[at.Value]; ch = true;
                if (e.Attribute(W + "initials") is { } ini) ini.Value = "A";
            }
            if (ppl)
                foreach (var per in x.Descendants(W15 + "person").ToList())
                {
                    var a = per.Attribute(W15 + "author")?.Value;
                    if (a != null && map.TryGetValue(a, out var nn) && sel.Any(s => s.Kind == PrivacyKind.Person && s.Value == a)) { per.Attribute(W15 + "author")!.Value = nn; ch = true; }
                    var pres = per.Element(W15 + "presenceInfo");
                    if (pres != null && sel.Any(s => s.Kind == PrivacyKind.PresenceInfo && s.Value == pres.Attribute(W15 + "userId")?.Value)) { pres.Remove(); ch = true; }
                }
            if (ch) p[name] = Bytes(x);
        }
        foreach (var x in sel.Where(x => x.Kind is PrivacyKind.CommentAuthor or PrivacyKind.RevisionAuthor or PrivacyKind.Person))
            done.Add(x.Kind == PrivacyKind.CommentAuthor && comments == CommentMode.Remove ? $"Commenti di «{x.Value}» rimossi" : $"«{x.Value}» → «{map[x.Value]}»");
        foreach (var x in sel.Where(x => x.Kind == PrivacyKind.PresenceInfo)) done.Add("Account dell'autore rimosso da people.xml");

        // ---- percorsi locali
        foreach (var f in sel.Where(x => x.Kind == PrivacyKind.LocalPath))
        {
            if (!p.TryGetValue(f.Part, out var b)) continue;
            var t = System.Text.Encoding.UTF8.GetString(b);
            var t2 = PathRx.Replace(t, m => m.Groups["u"].Success ? m.Value.Replace(m.Groups["u"].Value, "utente") : m.Value.Replace(m.Groups["srv"].Value, "server"));
            if (t2 != t) { p[f.Part] = System.Text.Encoding.UTF8.GetBytes(t2); done.Add($"Percorso in {f.Part} reso anonimo"); }
        }

        // ---- miniatura
        if (sel.Any(x => x.Kind == PrivacyKind.Thumbnail))
        {
            foreach (var k in p.Keys.Where(k => k.StartsWith("docProps/thumbnail", StringComparison.OrdinalIgnoreCase)).ToList()) { p.Remove(k); order.Remove(k); }
            if (Xml(p, "_rels/.rels") is { } rels)
            {
                rels.Descendants().Where(e => e.Name.LocalName == "Relationship" && (e.Attribute("Type")?.Value ?? "").EndsWith("/thumbnail", StringComparison.OrdinalIgnoreCase)).Remove();
                p["_rels/.rels"] = Bytes(rels);
            }
            done.Add("Anteprima miniatura rimossa");
        }

        // ---- riscrittura (stesso ordine delle parti, tutto il resto invariato byte per byte)
        using var ms = new MemoryStream();
        using (var z = new ZipArchive(ms, ZipArchiveMode.Create, true))
            foreach (var n in order.Where(p.ContainsKey))
            {
                var e = z.CreateEntry(n, CompressionLevel.Optimal);
                using var s = e.Open(); s.Write(p[n]);
            }
        var copy = ms.ToArray();
        return new CleanResult(copy, done, Scan(copy));
    }
    static readonly XNamespace W16cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
}
