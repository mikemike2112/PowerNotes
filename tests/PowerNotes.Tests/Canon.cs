using System.Xml.Linq;

namespace PowerNotes.Tests;

/// <summary>Forma canonica semantica di un part XML: attributi ordinati, dichiarazioni xmlns ignorate, whitespace tra elementi ignorato.</summary>
static class Canon
{
    public static string Of(Stream s, bool unordered = false)
    {
        var d = XDocument.Load(s, LoadOptions.PreserveWhitespace);
        var sb = new System.Text.StringBuilder();
        if (unordered)
        {
            // [Content_Types].xml e .rels sono insiemi: l'ordine non è semantico
            return string.Join("|", d.Root!.Elements().Select(c => { var b = new System.Text.StringBuilder(); Write(c, b); return b.ToString(); }).OrderBy(x => x, StringComparer.Ordinal));
        }
        Write(d.Root!, sb);
        return sb.ToString();
    }

    static void Write(XElement e, System.Text.StringBuilder sb)
    {
        sb.Append('<').Append(e.Name.NamespaceName).Append('|').Append(e.Name.LocalName);
        foreach (var a in e.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.NamespaceName + "|" + a.Name.LocalName, StringComparer.Ordinal))
            sb.Append(' ').Append(a.Name.NamespaceName).Append('|').Append(a.Name.LocalName).Append('=').Append(a.Value);
        sb.Append('>');
        foreach (var n in e.Nodes())
        {
            if (n is XElement c) Write(c, sb);
            else if (n is XText t && (!string.IsNullOrWhiteSpace(t.Value) || e.Name.LocalName is "t" or "instrText" or "delText")) sb.Append(t.Value);
        }
        sb.Append("</>");
    }
}
