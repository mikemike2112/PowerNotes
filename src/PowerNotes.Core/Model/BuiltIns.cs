using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace PowerNotes.Core.Model;

/// <summary>Stili e definizioni di numerazione predefiniti, aggiunti al documento solo se mancanti.</summary>
public static class BuiltIns
{
    const string W = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    public record StyleInfo(string Id, string Name, string WordName);

    /// <summary>Stili proposti dalla galleria (id, nome visibile).</summary>
    public static readonly StyleInfo[] Gallery =
    [
        new("Normal", "Normale", "Normal"), new("Title", "Titolo", "Title"), new("Subtitle", "Sottotitolo", "Subtitle"),
        new("Heading1", "Titolo 1", "heading 1"), new("Heading2", "Titolo 2", "heading 2"), new("Heading3", "Titolo 3", "heading 3"),
        new("Quote", "Citazione", "Quote"), new("ListParagraph", "Paragrafo elenco", "List Paragraph"),
    ];

    static readonly Dictionary<string, string> StyleXml = new()
    {
        ["Normal"] = $"<w:style {W} w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style>",
        ["Title"] = $"<w:style {W} w:type=\"paragraph\" w:styleId=\"Title\"><w:name w:val=\"Title\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:uiPriority w:val=\"10\"/><w:qFormat/><w:pPr><w:spacing w:after=\"80\" w:line=\"240\" w:lineRule=\"auto\"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii=\"Calibri Light\" w:hAnsi=\"Calibri Light\"/><w:spacing w:val=\"-10\"/><w:kern w:val=\"28\"/><w:sz w:val=\"56\"/><w:szCs w:val=\"56\"/></w:rPr></w:style>",
        ["Subtitle"] = $"<w:style {W} w:type=\"paragraph\" w:styleId=\"Subtitle\"><w:name w:val=\"Subtitle\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:uiPriority w:val=\"11\"/><w:qFormat/><w:rPr><w:color w:val=\"595959\"/><w:spacing w:val=\"15\"/><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/></w:rPr></w:style>",
        ["Heading1"] = $"<w:style {W} w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:uiPriority w:val=\"9\"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before=\"240\" w:after=\"0\"/><w:outlineLvl w:val=\"0\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Calibri Light\" w:hAnsi=\"Calibri Light\"/><w:color w:val=\"2F5496\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:style>",
        ["Heading2"] = $"<w:style {W} w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"heading 2\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:uiPriority w:val=\"9\"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before=\"40\" w:after=\"0\"/><w:outlineLvl w:val=\"1\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Calibri Light\" w:hAnsi=\"Calibri Light\"/><w:color w:val=\"2F5496\"/><w:sz w:val=\"26\"/><w:szCs w:val=\"26\"/></w:rPr></w:style>",
        ["Heading3"] = $"<w:style {W} w:type=\"paragraph\" w:styleId=\"Heading3\"><w:name w:val=\"heading 3\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:uiPriority w:val=\"9\"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before=\"40\" w:after=\"0\"/><w:outlineLvl w:val=\"2\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Calibri Light\" w:hAnsi=\"Calibri Light\"/><w:color w:val=\"1F3763\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr></w:style>",
        ["Quote"] = $"<w:style {W} w:type=\"paragraph\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:uiPriority w:val=\"29\"/><w:qFormat/><w:pPr><w:spacing w:before=\"200\" w:after=\"160\"/><w:ind w:left=\"864\" w:right=\"864\"/><w:jc w:val=\"center\"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val=\"404040\"/></w:rPr></w:style>",
        ["ListParagraph"] = $"<w:style {W} w:type=\"paragraph\" w:styleId=\"ListParagraph\"><w:name w:val=\"List Paragraph\"/><w:basedOn w:val=\"Normal\"/><w:uiPriority w:val=\"34\"/><w:qFormat/><w:pPr><w:ind w:left=\"720\"/><w:contextualSpacing/></w:pPr></w:style>",
        ["TableGrid"] = $"<w:style {W} w:type=\"table\" w:styleId=\"TableGrid\"><w:name w:val=\"Table Grid\"/><w:basedOn w:val=\"TableNormal\"/><w:uiPriority w:val=\"39\"/><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/></w:tblBorders></w:tblPr></w:style>",
        ["TableNormal"] = $"<w:style {W} w:type=\"table\" w:default=\"1\" w:styleId=\"TableNormal\"><w:name w:val=\"Normal Table\"/><w:uiPriority w:val=\"99\"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w=\"0\" w:type=\"dxa\"/><w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style>",
        ["FootnoteText"] = $"<w:style {W} w:type=\"paragraph\" w:styleId=\"FootnoteText\"><w:name w:val=\"footnote text\"/><w:basedOn w:val=\"Normal\"/><w:uiPriority w:val=\"99\"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr><w:rPr><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:style>",
        ["FootnoteReference"] = $"<w:style {W} w:type=\"character\" w:styleId=\"FootnoteReference\"><w:name w:val=\"footnote reference\"/><w:uiPriority w:val=\"99\"/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr></w:style>",
        ["Header"] = $"<w:style {W} w:type=\"paragraph\" w:styleId=\"Header\"><w:name w:val=\"header\"/><w:basedOn w:val=\"Normal\"/><w:uiPriority w:val=\"99\"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val=\"center\" w:pos=\"4513\"/><w:tab w:val=\"right\" w:pos=\"9026\"/></w:tabs><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr></w:style>",
        ["Footer"] = $"<w:style {W} w:type=\"paragraph\" w:styleId=\"Footer\"><w:name w:val=\"footer\"/><w:basedOn w:val=\"Normal\"/><w:uiPriority w:val=\"99\"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val=\"center\" w:pos=\"4513\"/><w:tab w:val=\"right\" w:pos=\"9026\"/></w:tabs><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr></w:style>",
    };

    public static Styles NewStyles() => new($"<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\" w:eastAsia=\"Calibri\" w:cs=\"Calibri\"/><w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/><w:lang w:val=\"it-IT\" w:eastAsia=\"en-US\" w:bidi=\"ar-SA\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after=\"160\" w:line=\"259\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults>{StyleXml["Normal"].Replace(" " + W, "")}{StyleXml["TableNormal"].Replace(" " + W, "")}</w:styles>");

    /// <summary>Garantisce che lo stile esista (cercandolo per nome: gli id sono localizzati) e ne restituisce l'id reale.</summary>
    public static string EnsureStyle(WordprocessingDocument doc, string key)
    {
        var main = doc.MainDocumentPart!;
        var sp = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        sp.Styles ??= NewStyles();
        if (!StyleXml.TryGetValue(key, out var xml)) return key;
        var name = System.Text.RegularExpressions.Regex.Match(xml, "<w:name w:val=\"([^\"]*)\"").Groups[1].Value;
        var found = sp.Styles.Elements<Style>().FirstOrDefault(s => string.Equals(s.StyleName?.Val?.Value, name, StringComparison.OrdinalIgnoreCase));
        if (found?.StyleId?.Value is { } fid) return fid;
        string normal = key == "Normal" ? "Normal" : EnsureStyle(doc, "Normal");
        string tnormal = key is "TableNormal" or "Normal" ? "TableNormal" : (xml.Contains("TableNormal") ? EnsureStyle(doc, "TableNormal") : "TableNormal");
        xml = xml.Replace("w:basedOn w:val=\"Normal\"", $"w:basedOn w:val=\"{normal}\"").Replace("w:next w:val=\"Normal\"", $"w:next w:val=\"{normal}\"").Replace("w:basedOn w:val=\"TableNormal\"", $"w:basedOn w:val=\"{tnormal}\"");
        sp.Styles.Append(new Style(xml));
        return key;
    }

    static string Lvl(int i, string fmt, string text, int left, string? font = null) =>
        $"<w:lvl w:ilvl=\"{i}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{fmt}\"/><w:lvlText w:val=\"{text}\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"{left}\" w:hanging=\"360\"/></w:pPr>{(font == null ? "" : $"<w:rPr><w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\" w:hint=\"default\"/></w:rPr>")}</w:lvl>";

    /// <summary>Crea (se manca) una definizione di elenco e restituisce il numId. kind: bullet | decimal | multilevel.</summary>
    public static int EnsureList(WordprocessingDocument doc, string kind)
    {
        var main = doc.MainDocumentPart!;
        var np = main.NumberingDefinitionsPart ?? main.AddNewPart<NumberingDefinitionsPart>();
        np.Numbering ??= new Numbering($"<w:numbering {W}/>");
        var num = np.Numbering;
        var existing = num.Elements<AbstractNum>().FirstOrDefault(a => a.Nsid?.Val?.Value == (kind == "bullet" ? "504E0001" : kind == "decimal" ? "504E0002" : "504E0003"));
        int absId;
        if (existing != null) absId = existing.AbstractNumberId!.Value;
        else
        {
            absId = num.Elements<AbstractNum>().Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(-1).Max() + 1;
            var sb = new System.Text.StringBuilder();
            string nsid = kind == "bullet" ? "504E0001" : kind == "decimal" ? "504E0002" : "504E0003";
            sb.Append($"<w:abstractNum {W} w:abstractNumId=\"{absId}\"><w:nsid w:val=\"{nsid}\"/><w:multiLevelType w:val=\"hybridMultilevel\"/>");
            for (int i = 0; i < 9; i++)
            {
                int left = 720 * (i + 1);
                if (kind == "bullet") sb.Append(Lvl(i, "bullet", i % 3 == 0 ? "" : i % 3 == 1 ? "o" : "", left, i % 3 == 0 ? "Symbol" : i % 3 == 1 ? "Courier New" : "Wingdings"));
                else if (kind == "decimal") sb.Append(Lvl(i, i % 3 == 0 ? "decimal" : i % 3 == 1 ? "lowerLetter" : "lowerRoman", $"%{i + 1}{(i % 3 == 2 ? ")" : ".")}", left));
                else sb.Append(Lvl(i, "decimal", string.Join(".", Enumerable.Range(1, i + 1).Select(k => "%" + k)) + ".", left + 0));
            }
            sb.Append("</w:abstractNum>");
            var abs = new AbstractNum(sb.ToString());
            var firstNum = num.Elements<NumberingInstance>().FirstOrDefault();
            if (firstNum != null) num.InsertBefore(abs, firstNum); else num.Append(abs);
        }
        var inst = num.Elements<NumberingInstance>().FirstOrDefault(n => n.AbstractNumId?.Val?.Value == absId);
        if (inst != null) return inst.NumberID!.Value;
        int numId = num.Elements<NumberingInstance>().Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
        num.Append(new NumberingInstance($"<w:num {W} w:numId=\"{numId}\"><w:abstractNumId w:val=\"{absId}\"/></w:num>"));
        return numId;
    }

    /// <summary>Tipo di elenco (bullet/decimal/multilevel) se la definizione è una delle nostre, altrimenti dedotto dal formato del livello 0.</summary>
    public static string? KindOf(WordprocessingDocument doc, int? numId)
    {
        if (numId is null or 0) return null;
        var num = doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        var inst = num?.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numId);
        var abs = num?.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == inst?.AbstractNumId?.Val?.Value);
        if (abs == null) return null;
        return abs.Nsid?.Val?.Value switch { "504E0001" => "bullet", "504E0002" => "decimal", "504E0003" => "multilevel", _ => abs.Elements<Level>().FirstOrDefault()?.NumberingFormat?.Val?.InnerText == "bullet" ? "bullet" : "decimal" };
    }

    /// <summary>Nuova istanza della stessa definizione, con ripartenza da 1 (elenco separato).</summary>
    public static int NewInstance(WordprocessingDocument doc, int baseNumId)
    {
        var num = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var inst = num.Elements<NumberingInstance>().First(n => n.NumberID?.Value == baseNumId);
        int numId = num.Elements<NumberingInstance>().Select(n => n.NumberID?.Value ?? 0).Max() + 1;
        num.Append(new NumberingInstance($"<w:num {W} w:numId=\"{numId}\"><w:abstractNumId w:val=\"{inst.AbstractNumId!.Val!.Value}\"/><w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"1\"/></w:lvlOverride></w:num>"));
        return numId;
    }
}
