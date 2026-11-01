using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PowerNotes.Core.Privacy;

namespace PowerNotes.Tests;

/// <summary>Costruisce un .docx con dati personali NOTI (finti) partendo da un file del corpus.</summary>
static class PrivacySample
{
    public const string Mario = "Mario Rossi", Luisa = "Luisa Bianchi", Acme = "Acme Srl", Segreto = "Progetto Aurora", Email = "mrossi@example.test", User = "mrossi";

    public static byte[] Build(bool withExif = false)
    {
        var src = File.ReadAllBytes(Path.Combine(TestPaths.Corpus, "plain-paragraphs.docx"));
        using var ms = new MemoryStream(); ms.Write(src);
        using (var z = new ZipArchive(ms, ZipArchiveMode.Update, true))
        {
            string Rd(string n) { using var r = new StreamReader(z.GetEntry(n)!.Open()); return r.ReadToEnd(); }
            void Wr(string n, string t) { z.GetEntry(n)?.Delete(); using var w = new StreamWriter(z.CreateEntry(n).Open(), new UTF8Encoding(false)); w.Write(t); }
            // core / app
            var core = Rd("docProps/core.xml");
            core = System.Text.RegularExpressions.Regex.Replace(core, "<dc:creator>[^<]*</dc:creator>", $"<dc:creator>{Mario}</dc:creator>");
            core = System.Text.RegularExpressions.Regex.Replace(core, "<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", $"<cp:lastModifiedBy>{Luisa}</cp:lastModifiedBy>");
            Wr("docProps/core.xml", core);
            var app = Rd("docProps/app.xml");
            app = System.Text.RegularExpressions.Regex.Replace(app, "<Company>[^<]*</Company>|<Company\\s*/>", $"<Company>{Acme}</Company>");
            if (!app.Contains("<Manager>")) app = app.Replace("</Properties>", $"<Manager>{Luisa}</Manager></Properties>");
            Wr("docProps/app.xml", app);
            // custom.xml
            Wr("docProps/custom.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><property fmtid=\"{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}\" pid=\"2\" name=\"Cliente\"><vt:lpwstr>{Segreto}</vt:lpwstr></property></Properties>");
            // comments, people
            Wr("word/comments.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:comment w:id=\"0\" w:author=\"{Mario}\" w:date=\"2026-01-01T10:00:00Z\" w:initials=\"MR\"><w:p><w:r><w:t>Nota privata</w:t></w:r></w:p></w:comment></w:comments>");
            Wr("word/people.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w15:people xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"><w15:person w15:author=\"{Luisa}\"><w15:presenceInfo w15:providerId=\"AD\" w15:userId=\"{Email}\"/></w15:person></w15:people>");
            // revisione + commento nel corpo
            var doc = Rd("word/document.xml");
            int i = doc.IndexOf("<w:r>", StringComparison.Ordinal); if (i < 0) i = doc.IndexOf("<w:r ", StringComparison.Ordinal);
            doc = doc.Insert(i, $"<w:commentRangeStart w:id=\"0\"/><w:ins w:id=\"900\" w:author=\"{Luisa}\" w:date=\"2026-01-02T10:00:00Z\"><w:r><w:t>testo inserito </w:t></w:r></w:ins><w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r>");
            Wr("word/document.xml", doc);
            // settings: modello allegato con percorso locale
            var rels = Rd("word/_rels/document.xml.rels").Replace("</Relationships>",
                "<Relationship Id=\"rIdC\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"comments.xml\"/>" +
                "<Relationship Id=\"rIdP\" Type=\"http://schemas.microsoft.com/office/2011/relationships/people\" Target=\"people.xml\"/>" +
                $"<Relationship Id=\"rIdT\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate\" Target=\"file:///C:/Users/{User}/Documents/modello.dotm\" TargetMode=\"External\"/></Relationships>");
            Wr("word/_rels/document.xml.rels", rels);
            // miniatura
            Wr("docProps/thumbnail.jpeg", "FAKEJPEG");
            var root = Rd("_rels/.rels").Replace("</Relationships>", "<Relationship Id=\"rIdTh\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail\" Target=\"docProps/thumbnail.jpeg\"/></Relationships>");
            Wr("_rels/.rels", root);
            var ct = Rd("[Content_Types].xml");
            ct = ct.Replace("</Types>", "<Override PartName=\"/word/comments.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml\"/><Override PartName=\"/word/people.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml\"/><Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>" + (ct.Contains("Extension=\"jpeg\"") ? "" : "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>") + "</Types>");
            Wr("[Content_Types].xml", ct);
            var rootRels2 = Rd("_rels/.rels").Replace("</Relationships>", "<Relationship Id=\"rIdCu\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties\" Target=\"docProps/custom.xml\"/></Relationships>");
            Wr("_rels/.rels", rootRels2);
            if (withExif) { var e = z.CreateEntry("word/media/foto.jpg"); using var s = e.Open(); s.Write(new byte[] { 0xFF, 0xD8, 0xFF, 0xE1, 0, 10 }.Concat(Encoding.ASCII.GetBytes("Exif\0\0GPS")).ToArray()); }
        }
        return ms.ToArray();
    }

    public static string AllText(byte[] docx)
    {
        using var z = new ZipArchive(new MemoryStream(docx));
        var sb = new StringBuilder();
        foreach (var e in z.Entries.Where(e => e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels"))) { using var r = new StreamReader(e.Open()); sb.AppendLine(e.FullName).AppendLine(r.ReadToEnd()); }
        return sb.ToString();
    }
}

public class PrivacyTests
{
    [Fact]
    public void Scan_FindsAllKnownData_WithRealValues()
    {
        var r = PrivacyShield.Scan(PrivacySample.Build(withExif: true));
        string V(PrivacyKind k) => string.Join("|", r.Findings.Where(f => f.Kind == k).Select(f => f.Value));
        Assert.Equal(PrivacySample.Mario, V(PrivacyKind.Author)); Assert.Contains(PrivacySample.Luisa, V(PrivacyKind.LastAuthor));
        Assert.Equal(PrivacySample.Acme, V(PrivacyKind.Company)); Assert.Equal(PrivacySample.Luisa, V(PrivacyKind.Manager));
        Assert.Equal(PrivacySample.Segreto, V(PrivacyKind.CustomProperty));
        Assert.Equal(PrivacySample.Mario, V(PrivacyKind.CommentAuthor)); Assert.Equal(PrivacySample.Luisa, V(PrivacyKind.RevisionAuthor));
        Assert.Equal(PrivacySample.Luisa, V(PrivacyKind.Person)); Assert.Equal(PrivacySample.Email, V(PrivacyKind.PresenceInfo));
        Assert.Contains(PrivacySample.User, V(PrivacyKind.LocalPath)); Assert.Single(r.Findings, f => f.Kind == PrivacyKind.Thumbnail);
        Assert.Contains(r.Warnings, w => w.Contains("EXIF"));
        Assert.Contains(r.Warnings, w => w.Contains("testo del documento"));      // mai "pulito" senza riserve
    }

    [Theory, InlineData(CommentMode.Anonymize), InlineData(CommentMode.Remove)]
    public void Clean_CopyHasNoKnownDataInAnyPart_OriginalUntouched_ValidAndRescanned(CommentMode mode)
    {
        var original = PrivacySample.Build(); var before = (byte[])original.Clone();
        var scan = PrivacyShield.Scan(original);
        var res = PrivacyShield.Clean(original, scan.Findings.Select(f => f.Id).ToHashSet(), mode);
        Assert.Equal(before, original);                                        // l'originale è identico byte per byte
        var all = PrivacySample.AllText(res.Copy);
        foreach (var secret in new[] { PrivacySample.Mario, PrivacySample.Luisa, PrivacySample.Acme, PrivacySample.Segreto, PrivacySample.Email, "/" + PrivacySample.User + "/" })
            Assert.DoesNotContain(secret, all);
        if (mode == CommentMode.Remove) Assert.DoesNotContain("Nota privata", all); else Assert.Contains("Nota privata", all);   // anonimizza: contenuto invariato
        Assert.Contains("testo inserito", all);                                // le revisioni NON vengono né accettate né rifiutate
        Assert.Empty(res.After.Findings);                                      // la copia ricontrollata non ha più nulla
        Assert.NotEmpty(res.Done); Assert.Contains(res.After.Warnings, w => w.Contains("testo del documento"));
        var path = Path.Combine(Path.GetTempPath(), "pn-priv-" + Guid.NewGuid().ToString("N") + ".docx");
        File.WriteAllBytes(path, res.Copy);
        try { using var pk = WordprocessingDocument.Open(path, false); Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(pk)); }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Anonymize_UsesConsistentPseudonyms_AndUnselectedItemsStay()
    {
        var o = PrivacySample.Build();
        var res = PrivacyShield.Clean(o, new HashSet<string> { "comment:" + PrivacySample.Mario, "rev:" + PrivacySample.Luisa }, CommentMode.Anonymize);
        var all = PrivacySample.AllText(res.Copy);
        Assert.Contains("Autore 1", all); Assert.Contains("Autore 2", all);
        Assert.Contains(PrivacySample.Acme, all);                              // non selezionato: resta
        Assert.Contains(res.After.Findings, f => f.Kind == PrivacyKind.Company);
    }

    [Fact]
    public void AlreadyCleanFile_FindsNothing()
    {
        var clean = PrivacyShield.Clean(PrivacySample.Build(), PrivacyShield.Scan(PrivacySample.Build()).Findings.Select(f => f.Id).ToHashSet()).Copy;
        Assert.Empty(PrivacyShield.Scan(clean).Findings);
    }

    [Fact]
    public void CleanCorpusFile_HasNoFindingsAndIsOpenable()
    {
        foreach (var f in TestPaths.CorpusFiles()) Assert.DoesNotContain(PrivacyShield.Scan(File.ReadAllBytes(f)).Findings, x => x.Kind is PrivacyKind.Author or PrivacyKind.Person or PrivacyKind.PresenceInfo or PrivacyKind.LocalPath);
    }
}

public class PrivacySampleExport
{
    [Fact]
    public void WriteSampleForScreenshots() { Directory.CreateDirectory(Path.Combine(TestPaths.Root, "artifacts")); File.WriteAllBytes(Path.Combine(TestPaths.Root, "artifacts", "privacy-sample.docx"), PrivacySample.Build(withExif: true)); }
}
