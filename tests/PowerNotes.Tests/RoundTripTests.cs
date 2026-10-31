using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Model;

namespace PowerNotes.Tests;

public class RoundTripTests
{
    public static IEnumerable<object[]> Names => TestPaths.CorpusNames();

    static string Out(string name, string suffix = "")
    {
        var dir = Path.Combine(TestPaths.Root, "artifacts", "roundtrip");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, name + suffix + ".docx");
    }

    static List<string> Errors(string path)
    {
        using var d = WordprocessingDocument.Open(path, false);
        return new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(d).Select(e => e.ErrorType + ":" + e.Description).ToList();
    }

    static Dictionary<string, string> Parts(string path)
    {
        using var z = ZipFile.OpenRead(path);
        var res = new Dictionary<string, string>();
        foreach (var e in z.Entries.Where(e => !e.FullName.EndsWith('/')))
        {
            if (e.FullName == "[Content_Types].xml") continue;   // confrontato come mappa parte→tipo (vedi sotto)
            if (e.FullName.EndsWith(".rels")) { res[e.FullName] = "R" + Rels(e); continue; }
            using var s = e.Open();
            if (e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels") || e.FullName.EndsWith(".vml"))
            {
                try { res[e.FullName] = "X" + Canon.Of(s, e.FullName.EndsWith(".rels") || e.FullName == "[Content_Types].xml"); continue; } catch (System.Xml.XmlException) { }
            }
            using var ms = new MemoryStream();
            using var s2 = e.Open(); s2.CopyTo(ms);
            res[e.FullName] = "B" + Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(ms.ToArray()));
        }
        foreach (var kv in ContentTypes(z)) res["CT:" + kv.Key] = kv.Value;
        return res;
    }

    /// <summary>Relazioni come insieme di (Id, Type, Mode, destinazione risolta in nome di parte assoluto).</summary>
    static string Rels(ZipArchiveEntry e)
    {
        var doc = System.Xml.Linq.XDocument.Load(e.Open());
        var srcDir = e.FullName.Contains("_rels/") ? e.FullName[..e.FullName.IndexOf("_rels/", StringComparison.Ordinal)] : "";
        return string.Join("|", doc.Root!.Elements().Select(r =>
        {
            var t = r.Attribute("Target")!.Value; var mode = r.Attribute("TargetMode")?.Value ?? "Internal";
            if (mode == "Internal") t = (t.StartsWith('/') ? t : "/" + new Uri(new Uri("file:///" + srcDir), t).AbsolutePath.TrimStart('/')).Replace("%20", " ");
            return $"{r.Attribute("Id")!.Value};{r.Attribute("Type")!.Value};{mode};{t}";
        }).OrderBy(x => x, StringComparer.Ordinal));
    }

    /// <summary>Tipo di contenuto effettivo di ogni parte (Default per estensione, Override per nome): l'XML di [Content_Types] può cambiare forma senza cambiare significato.</summary>
    static Dictionary<string, string> ContentTypes(ZipArchive z)
    {
        var doc = System.Xml.Linq.XDocument.Load(z.GetEntry("[Content_Types].xml")!.Open());
        var def = doc.Root!.Elements().Where(x => x.Name.LocalName == "Default").ToDictionary(x => x.Attribute("Extension")!.Value.ToLowerInvariant(), x => x.Attribute("ContentType")!.Value);
        var ovr = doc.Root!.Elements().Where(x => x.Name.LocalName == "Override").ToDictionary(x => x.Attribute("PartName")!.Value.ToLowerInvariant(), x => x.Attribute("ContentType")!.Value);
        var res = new Dictionary<string, string>();
        foreach (var e in z.Entries.Where(e => !e.FullName.EndsWith('/') && e.FullName != "[Content_Types].xml"))
        {
            var name = "/" + e.FullName;
            var ext = Path.GetExtension(e.FullName).TrimStart('.').ToLowerInvariant();
            res[name] = ovr.TryGetValue(name.ToLowerInvariant(), out var t) ? t : def.TryGetValue(ext, out var d) ? d : "?";
        }
        return res;
    }

    [Theory, MemberData(nameof(Names))]
    public void OpenSaveNoChanges_IsValidAndPartsIdentical(string name)
    {
        var src = Path.Combine(TestPaths.Corpus, name + ".docx");
        var dst = Out(name);
        using (var d = WordDoc.Open(src)) { _ = d.Layout(); d.Save(dst); }

        // validatore: nessun errore nuovo rispetto all'originale
        var before = Errors(src).GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
        var after = Errors(dst).GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
        foreach (var kv in after) Assert.True(before.TryGetValue(kv.Key, out var n) && n >= kv.Value, "Errore nuovo dopo il salvataggio: " + kv.Key);

        // stesse parti, stesso contenuto XML canonicalizzato
        var a = Parts(src); var b = Parts(dst);
        Assert.Equal(a.Keys.OrderBy(x => x), b.Keys.OrderBy(x => x));
        foreach (var k in a.Keys) Assert.True(a[k] == b[k], $"Part modificata: {k}");
    }

    [Fact]
    public void OriginalValidatorErrors_AreReportedNotHidden()
    {
        // I falsi positivi del validatore sui file di Word si elencano (non si silenziano): vedi COMPAT.md
        var lines = new List<string>();
        foreach (var f in TestPaths.CorpusFiles()) { var e = Errors(f); if (e.Count > 0) lines.Add(Path.GetFileName(f) + ": " + string.Join(" | ", e.Distinct().Take(3))); }
        Directory.CreateDirectory(Path.Combine(TestPaths.Root, "artifacts"));
        File.WriteAllLines(Path.Combine(TestPaths.Root, "artifacts", "validator-original-errors.txt"), lines);
    }

    static string ReadEntry(ZipArchive z, string name) { using var r = new StreamReader(z.GetEntry(name)!.Open()); return r.ReadToEnd(); }
    static void WriteEntry(ZipArchive z, string name, string text)
    {
        z.GetEntry(name)?.Delete();
        using var w = new StreamWriter(z.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false));
        w.Write(text);
    }

    // ---- contenuto sconosciuto sopravvive a modifica e salvataggio ----
    [Fact]
    public void UnknownContent_SurvivesEditAndSave()
    {
        var src = Path.Combine(TestPaths.Corpus, "plain-paragraphs.docx");
        var tmp = Out("unknown-src");
        File.Copy(src, tmp, true);
        using (var z = ZipFile.Open(tmp, ZipArchiveMode.Update))
        {
            var xml = ReadEntry(z, "word/document.xml");
            xml = xml.Replace("<w:document ", "<w:document xmlns:zz=\"urn:test:zz\" ");
            xml = System.Text.RegularExpressions.Regex.Replace(xml, "mc:Ignorable=\"([^\"]*)\"", "mc:Ignorable=\"$1 zz\"");
            int i = xml.IndexOf("<w:p ", StringComparison.Ordinal); int j = xml.IndexOf('>', i);
            xml = xml.Insert(j, " zz:flag=\"1\"");
            int k = xml.IndexOf("<w:r>", i, StringComparison.Ordinal); if (k < 0) k = xml.IndexOf("<w:r ", i, StringComparison.Ordinal);
            xml = xml.Insert(k, "<zz:marker zz:id=\"42\"><zz:child/></zz:marker>");
            WriteEntry(z, "word/document.xml", xml);
            WriteEntry(z, "customXml/itemZZ.xml", "<root xmlns=\"urn:test:custom\"><v>keep-me</v></root>");
            WriteEntry(z, "word/_rels/document.xml.rels", ReadEntry(z, "word/_rels/document.xml.rels").Replace("</Relationships>", "<Relationship Id=\"rIdZZ\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/itemZZ.xml\"/></Relationships>"));
            WriteEntry(z, "[Content_Types].xml", ReadEntry(z, "[Content_Types].xml").Replace("</Types>", "<Override PartName=\"/customXml/itemZZ.xml\" ContentType=\"application/xml\"/></Types>"));
        }
        var dst = Out("unknown-dst");
        using (var d = WordDoc.Open(tmp))
        {
            var s = new EditorSession(d);
            s.Type("MODIFICA ");
            s.SelectAll();
            s.ApplyRun(r => r.Bold = new DocumentFormat.OpenXml.Wordprocessing.Bold());
            d.Save(dst);
        }
        using var z2 = ZipFile.OpenRead(dst);
        var doc = ReadEntry(z2, "word/document.xml");
        Assert.Contains("urn:test:zz", doc);
        Assert.Contains("zz:flag=\"1\"", doc);
        Assert.Contains("zz:marker", doc);
        Assert.Contains("zz:child", doc);
        Assert.Matches("mc:Ignorable=\"[^\"]*\\bzz\\b", doc);
        Assert.Contains("MODIFICA ", doc);
        Assert.NotNull(z2.GetEntry("customXml/itemZZ.xml"));
        Assert.Contains("itemZZ.xml", ReadEntry(z2, "word/_rels/document.xml.rels"));
    }

    [Fact]
    public void AtomicSave_FailureKeepsOriginalIntact()
    {
        var tmp = Out("atomic");
        File.Copy(Path.Combine(TestPaths.Corpus, "plain-paragraphs.docx"), tmp, true);
        var before = File.ReadAllBytes(tmp);
        using var d = WordDoc.Open(tmp);
        using (var lockFile = new FileStream(tmp, FileMode.Open, FileAccess.Read, FileShare.None))
            Assert.ThrowsAny<IOException>(() => d.Save(tmp));
        Assert.Equal(before, File.ReadAllBytes(tmp));
        Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(tmp)!, "atomic.docx.*.tmp"));
    }

    [Fact]
    public void MalformedFiles_GiveClearError()
    {
        var bad = Out("bad-garbage"); File.WriteAllText(bad, "non sono uno zip");
        Assert.Throws<DocumentLoadException>(() => WordDoc.Open(bad));
        var bad2 = Out("bad-xml"); File.Copy(Path.Combine(TestPaths.Corpus, "plain-paragraphs.docx"), bad2, true);
        using (var z = ZipFile.Open(bad2, ZipArchiveMode.Update)) WriteEntry(z, "word/document.xml", "<w:document><w:body>");
        Assert.Throws<DocumentLoadException>(() => { using var d = WordDoc.Open(bad2); _ = d.Layout(); });
    }
}
