using System.IO.Compression;
using System.Xml.Linq;

namespace PowerNotes.Tests;

/// <summary>Controlli deterministici sul package OPC: ogni relazione interna punta a una parte esistente, ogni r:id usato esiste, ogni parte ha un tipo.</summary>
static class PackageIntegrity
{
    const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static List<string> Problems(string path)
    {
        var res = new List<string>();
        try
        {
            using var z = ZipFile.OpenRead(path);
            var names = z.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (!names.Contains("[Content_Types].xml")) { res.Add("manca [Content_Types].xml"); return res; }
            var ct = XDocument.Load(z.GetEntry("[Content_Types].xml")!.Open());
            var defs = ct.Root!.Elements().Where(x => x.Name.LocalName == "Default").Select(x => x.Attribute("Extension")!.Value.ToLowerInvariant()).ToHashSet();
            var ovr = ct.Root!.Elements().Where(x => x.Name.LocalName == "Override").Select(x => x.Attribute("PartName")!.Value.ToLowerInvariant()).ToHashSet();
            foreach (var e in z.Entries.Where(e => !e.FullName.EndsWith('/') && e.FullName != "[Content_Types].xml"))
                if (!ovr.Contains("/" + e.FullName.ToLowerInvariant()) && !defs.Contains(Path.GetExtension(e.FullName).TrimStart('.').ToLowerInvariant())) res.Add("parte senza tipo: " + e.FullName);
            foreach (var e in z.Entries.Where(e => e.FullName.EndsWith(".rels")))
            {
                var rels = XDocument.Load(e.Open());
                var srcDir = e.FullName.Contains("_rels/") ? e.FullName[..e.FullName.IndexOf("_rels/", StringComparison.Ordinal)] : "";
                var ids = new HashSet<string>();
                foreach (var r in rels.Root!.Elements())
                {
                    ids.Add(r.Attribute("Id")!.Value);
                    if ((r.Attribute("TargetMode")?.Value ?? "Internal") != "Internal") continue;
                    var t = r.Attribute("Target")!.Value;
                    var abs = t.StartsWith('/') ? t.TrimStart('/') : new Uri(new Uri("file:///" + srcDir), t).AbsolutePath.TrimStart('/');
                    abs = Uri.UnescapeDataString(abs);
                    if (!names.Contains(abs)) res.Add($"{e.FullName}: destinazione mancante {t}");
                }
                var partName = srcDir + Path.GetFileName(e.FullName)[..^5];
                var part = z.GetEntry(partName);
                if (part == null || !partName.EndsWith(".xml")) continue;
                XDocument x;
                try { x = XDocument.Load(part.Open()); } catch (System.Xml.XmlException) { res.Add("XML rotto: " + partName); continue; }
                foreach (var a in x.Descendants().Attributes().Where(a => a.Name.NamespaceName == R && a.Name.LocalName is "id" or "embed" or "link" or "pict" or "href"))
                    if (!ids.Contains(a.Value)) res.Add($"{partName}: r:{a.Name.LocalName}={a.Value} senza relazione");
            }
        }
        catch (Exception e) when (e is InvalidDataException or IOException or System.Xml.XmlException) { res.Add("package illeggibile: " + e.Message); }
        return res;
    }
}
