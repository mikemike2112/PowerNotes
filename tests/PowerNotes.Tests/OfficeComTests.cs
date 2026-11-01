using System.Diagnostics;
using System.IO.Compression;
using System.Text.Json;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Model;
using PowerNotes.Core.Privacy;

namespace PowerNotes.Tests;

/// <summary>Esegue UNA sessione Word (COM, script in tools/) su originali, salvati, modificati e file corrotti; i test leggono il JSON.</summary>
public sealed class OfficeFixture : IDisposable
{
    public static readonly string[] ModSubset =
    [
        "plain-paragraphs", "text-formatting", "alignment", "list-bullets", "list-numbered", "table-simple", "table-merged-cells",
        "image-inline", "header-footer", "page-numbers", "section-break", "mixed-report", "mixed-illustrated", "tabs"
    ];
    public const string Marker = "MODIFICA ";

    public string Dir { get; } = Path.Combine(TestPaths.Root, "artifacts", "com");
    public Dictionary<string, JsonElement> Result { get; } = new(StringComparer.OrdinalIgnoreCase);
    public string[] Negative { get; } = ["neg-xml-rotto", "neg-tipo-contenuto", "neg-parte-mancante", "neg-relazione-orfana"];
    public string[] WordMustFail { get; } = ["neg-xml-rotto", "neg-tipo-contenuto"];
    JsonDocument? _json;

    public OfficeFixture()
    {
        if (Directory.Exists(Dir)) Directory.Delete(Dir, true);
        Directory.CreateDirectory(Dir);
        var files = new List<string>();
        foreach (var f in TestPaths.CorpusFiles())
        {
            var name = Path.GetFileNameWithoutExtension(f);
            files.Add(f);
            var saved = Path.Combine(Dir, name + ".saved.docx");
            using (var d = WordDoc.Open(f)) d.Save(saved);
            files.Add(saved);
            if (ModSubset.Contains(name))
            {
                var mod = Path.Combine(Dir, name + ".mod.docx");
                using var d = WordDoc.Open(f);
                var s = new EditorSession(d);
                s.Type(Marker);
                s.ReplaceAll("paragrafo", "capitolo", false);
                s.SelectAll(); s.ToggleBold();
                d.Save(mod);
                files.Add(mod);
            }
        }
        // Uniforma e Scudo privacy: i file prodotti devono aprirsi in Word senza ripristino e con lo stesso testo
        foreach (var n in new[] { "text-formatting", "mixed-report", "table-borders-shading" })
        {
            var u = Path.Combine(Dir, n + ".uniform.docx");
            using (var d = WordDoc.Open(Path.Combine(TestPaths.Corpus, n + ".docx"))) { new EditorSession(d).Uniform(); d.Save(u); }
            files.Add(u);
        }
        var sample = PrivacySample.Build(); var samplePath = Path.Combine(Dir, "privacy-sample.docx"); File.WriteAllBytes(samplePath, sample); files.Add(samplePath);
        var cleaned = PrivacyShield.Clean(sample, PrivacyShield.Scan(sample).Findings.Select(f => f.Id).ToHashSet()).Copy;
        File.WriteAllBytes(Path.Combine(Dir, "privacy-clean.docx"), cleaned); files.Add(Path.Combine(Dir, "privacy-clean.docx"));
        var cleanedRm = PrivacyShield.Clean(sample, PrivacyShield.Scan(sample).Findings.Select(f => f.Id).ToHashSet(), CommentMode.Remove).Copy;
        File.WriteAllBytes(Path.Combine(Dir, "privacy-clean-rm.docx"), cleanedRm); files.Add(Path.Combine(Dir, "privacy-clean-rm.docx"));
        // controlli negativi: il rilevatore DEVE fallire su questi
        var baseFile = Path.Combine(TestPaths.Corpus, "table-simple.docx");
        MakeBad("neg-xml-rotto", baseFile, z => Write(z, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>rotto</w:r></w:p></w:body>"), files);
        MakeBad("neg-tipo-contenuto", baseFile, z =>
        {
            string ct; using (var r = new StreamReader(z.GetEntry("[Content_Types].xml")!.Open())) ct = r.ReadToEnd();
            Write(z, "[Content_Types].xml", ct.Replace("wordprocessingml.document.main+xml", "wordprocessingml.styles+xml"));
        }, files);
        MakeBad("neg-parte-mancante", baseFile, z => z.GetEntry("word/styles.xml")!.Delete(), files);
        MakeBad("neg-relazione-orfana", baseFile, z => { }, files);
        // la relazione orfana è inutile senza un riferimento: aggiungilo al documento
        var orphan = Path.Combine(Dir, "neg-relazione-orfana.docx");
        using (var z = ZipFile.Open(orphan, ZipArchiveMode.Update))
        {
            string x; using (var r = new StreamReader(z.GetEntry("word/document.xml")!.Open())) x = r.ReadToEnd();
            x = x.Replace("<w:body>", "<w:body><w:p><w:r><w:drawing><wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"><wp:extent cx=\"914400\" cy=\"914400\"/><wp:docPr id=\"99\" name=\"x\"/><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><pic:nvPicPr><pic:cNvPr id=\"99\" name=\"x\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rId99\"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>");
            string rx; using (var r = new StreamReader(z.GetEntry("word/_rels/document.xml.rels")!.Open())) rx = r.ReadToEnd();
            Write(z, "word/_rels/document.xml.rels", rx.Replace("</Relationships>", "<Relationship Id=\"rId99\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/non-esiste.png\"/></Relationships>"));
            Write(z, "word/document.xml", x);
        }

        var json = Path.Combine(Dir, "result.json");
        var psi = new ProcessStartInfo("powershell", $"-NoProfile -ExecutionPolicy Bypass -File \"{Path.Combine(TestPaths.Tools, "word-inspect.ps1")}\" -Out \"{json}\" -Paths \"{string.Join(",", files)}\"")
        { UseShellExecute = false, RedirectStandardError = true, RedirectStandardOutput = true, CreateNoWindow = true };
        using var p = Process.Start(psi)!;
        var err = p.StandardError.ReadToEndAsync(); var outp = p.StandardOutput.ReadToEndAsync();
        if (!p.WaitForExit(900_000)) { try { p.Kill(true); } catch (InvalidOperationException) { } throw new TimeoutException("Sessione Word COM oltre il tempo limite."); }
        if (!File.Exists(json)) throw new InvalidOperationException("word-inspect non ha prodotto output: " + err.Result);
        _json = JsonDocument.Parse(File.ReadAllText(json));
        foreach (var kv in _json.RootElement.EnumerateObject()) Result[Path.GetFileName(kv.Name)] = kv.Value;
    }

    static void Write(ZipArchive z, string name, string text)
    {
        z.GetEntry(name)?.Delete();
        using var w = new StreamWriter(z.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false));
        w.Write(text);
    }

    void MakeBad(string name, string src, Action<ZipArchive> corrupt, List<string> files)
    {
        var dst = Path.Combine(Dir, name + ".docx");
        File.Copy(src, dst, true);
        using (var z = ZipFile.Open(dst, ZipArchiveMode.Update)) corrupt(z);
        files.Add(dst);
    }

    public void Dispose() => _json?.Dispose();
}

[CollectionDefinition("OfficeCom", DisableParallelization = true)]
public class OfficeCollection : ICollectionFixture<OfficeFixture> { }

[Collection("OfficeCom"), Trait("Category", "OfficeCom")]
public class OfficeComTests(OfficeFixture fx)
{
    public static IEnumerable<object[]> Names => TestPaths.CorpusNames();
    public static IEnumerable<object[]> ModNames => OfficeFixture.ModSubset.Select(n => new object[] { n });

    JsonElement R(string file) => fx.Result.TryGetValue(file, out var e) ? e : throw new KeyNotFoundException(file);

    [Theory, MemberData(nameof(Names))]
    public void OriginalAndSaved_OpenInWordWithoutRepair_SameContent(string name)
    {
        var o = R(name + ".docx"); var s = R(name + ".saved.docx");
        Assert.True(o.GetProperty("ok").GetBoolean(), "originale non si apre: " + o);
        Assert.True(s.GetProperty("ok").GetBoolean(), "salvato non si apre in Word: " + s.GetProperty("error"));
        Assert.Equal(o.GetProperty("text").GetString(), s.GetProperty("text").GetString());
        foreach (var k in new[] { "pages", "paragraphs", "tables", "inlineShapes", "shapes", "sections" })
            Assert.Equal(o.GetProperty(k).GetInt32(), s.GetProperty(k).GetInt32());
        // nessun dato personale nei metadati del corpus
        Assert.Equal("Corpus", o.GetProperty("author").GetString());
        Assert.Equal("Corpus", o.GetProperty("lastAuthor").GetString());
    }

    [Theory, MemberData(nameof(ModNames))]
    public void EditedFile_HasChangeAndKeepsRest(string name)
    {
        var o = R(name + ".docx"); var m = R(name + ".mod.docx");
        Assert.True(m.GetProperty("ok").GetBoolean(), "file modificato non si apre: " + m.GetProperty("error"));
        var expected = OfficeFixture.Marker + o.GetProperty("text").GetString()!.Replace("paragrafo", "capitolo", StringComparison.OrdinalIgnoreCase);
        Assert.Equal(expected, m.GetProperty("text").GetString());
        foreach (var k in new[] { "tables", "inlineShapes", "sections" })
            Assert.Equal(o.GetProperty(k).GetInt32(), m.GetProperty(k).GetInt32());
    }

    [Theory, InlineData("text-formatting"), InlineData("mixed-report"), InlineData("table-borders-shading")]
    public void Uniform_OpensInWordWithSameText(string name)
    {
        var o = R(name + ".docx"); var u = R(name + ".uniform.docx");
        Assert.True(u.GetProperty("ok").GetBoolean(), "file uniformato non si apre in Word: " + u.GetProperty("error"));
        Assert.Equal(o.GetProperty("text").GetString(), u.GetProperty("text").GetString());
        foreach (var k in new[] { "tables", "inlineShapes", "sections" }) Assert.Equal(o.GetProperty(k).GetInt32(), u.GetProperty(k).GetInt32());
    }

    [Theory, InlineData("privacy-clean"), InlineData("privacy-clean-rm")]
    public void PrivacyCopy_OpensInWordWithoutRepair_SameBodyText_AuthorsGone(string file)
    {
        var o = R("privacy-sample.docx"); var c = R(file + ".docx");
        Assert.True(o.GetProperty("ok").GetBoolean(), "campione non si apre: " + o.GetProperty("error"));
        Assert.True(c.GetProperty("ok").GetBoolean(), "copia pulita non si apre in Word: " + c.GetProperty("error"));
        Assert.Equal(o.GetProperty("text").GetString(), c.GetProperty("text").GetString());   // stesso contenuto (revisione inserita compresa)
        Assert.NotEqual(PrivacySample.Mario, c.GetProperty("author").GetString());
        Assert.DoesNotContain(PrivacySample.Luisa, c.GetProperty("lastAuthor").GetString());
    }

    [Fact]
    public void NegativeControl_DetectorFailsOnCorruptedFiles()
    {
        // rilevatore = Word (apertura senza errore) + integrità del package (relazioni/parti/tipi). Word tollera in silenzio alcuni danni: vedi COMPAT.md
        foreach (var n in fx.Negative)
        {
            var broken = !R(n + ".docx").GetProperty("ok").GetBoolean() || PackageIntegrity.Problems(Path.Combine(fx.Dir, n + ".docx")).Count > 0;
            Assert.True(broken, $"Il rilevatore NON ha segnalato il file corrotto {n}: il risultato dei round-trip non varrebbe.");
        }
        foreach (var n in fx.WordMustFail)
            Assert.False(R(n + ".docx").GetProperty("ok").GetBoolean(), $"Word avrebbe dovuto rifiutare {n}.");
    }

    [Theory, MemberData(nameof(Names))]
    public void SavedFile_PackageIntegrityIsClean(string name)
    {
        Assert.Empty(PackageIntegrity.Problems(Path.Combine(fx.Dir, name + ".saved.docx")));
        Assert.Empty(PackageIntegrity.Problems(Path.Combine(TestPaths.Corpus, name + ".docx")));   // l'originale è pulito: il rilevatore non dà falsi positivi
    }
}
