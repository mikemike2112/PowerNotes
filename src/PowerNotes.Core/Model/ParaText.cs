using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace PowerNotes.Core.Model;

public enum LeafKind { Text, Tab, Break, PageBreak, Image }

/// <summary>Foglia inline di un paragrafo: testo, tab, a capo, immagine. Gli offset sono nello spazio dei caratteri del paragrafo.</summary>
public sealed class Leaf
{
    public required Run Run;
    public required OpenXmlElement El;      // Text, TabChar, Break, Drawing
    public LeafKind Kind;
    public string Text = "";
    public int Start, Length;
    public string? Field;                   // istruzione del campo di cui è risultato (es. "PAGE")
    public int End => Start + Length;
}

public static class ParaText
{
    /// <summary>Foglie in ordine di documento. Salta testo di caselle di testo annidate, contenuto cancellato (revisioni) e Fallback di mc:AlternateContent.</summary>
    public static List<Leaf> Leaves(Paragraph p)
    {
        var res = new List<Leaf>();
        int pos = 0;
        var fieldStack = new Stack<(System.Text.StringBuilder Instr, bool Result)>();
        foreach (var run in p.Descendants<Run>())
        {
            if (run.Ancestors<Paragraph>().First() != p) continue;
            if (run.Ancestors<DeletedRun>().Any()) continue;
            if (run.Ancestors<AlternateContentFallback>().Any()) continue;
            string? simple = run.Ancestors<SimpleField>().FirstOrDefault()?.Instruction?.Value?.Trim();
            foreach (var el in run.ChildElements)
            {
                switch (el)
                {
                    case FieldChar fc:
                        var t = fc.FieldCharType?.Value;
                        if (t == FieldCharValues.Begin) fieldStack.Push((new(), false));
                        else if (t == FieldCharValues.Separate && fieldStack.Count > 0) { var f = fieldStack.Pop(); fieldStack.Push((f.Instr, true)); }
                        else if (t == FieldCharValues.End && fieldStack.Count > 0) fieldStack.Pop();
                        continue;
                    case FieldCode fcode:
                        if (fieldStack.Count > 0) fieldStack.Peek().Instr.Append(fcode.Text);
                        continue;
                }
                if (fieldStack.Count > 0 && !fieldStack.Peek().Result) continue;   // istruzione nascosta
                string? field = fieldStack.Count > 0 ? fieldStack.Peek().Instr.ToString().Trim() : simple;
                Leaf? leaf = el switch
                {
                    Text tx => new Leaf { Run = run, El = el, Kind = LeafKind.Text, Text = tx.Text, Length = tx.Text.Length },
                    NoBreakHyphen => new Leaf { Run = run, El = el, Kind = LeafKind.Text, Text = "-", Length = 1 },
                    TabChar => new Leaf { Run = run, El = el, Kind = LeafKind.Tab, Text = "\t", Length = 1 },
                    CarriageReturn => new Leaf { Run = run, El = el, Kind = LeafKind.Break, Text = "\n", Length = 1 },
                    Break br => new Leaf { Run = run, El = el, Kind = br.Type?.Value == BreakValues.Page ? LeafKind.PageBreak : LeafKind.Break, Text = "\n", Length = 1 },
                    Drawing d when d.Descendants<A.Blip>().Any() && d.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>() != null
                        => new Leaf { Run = run, El = el, Kind = LeafKind.Image, Text = "￼", Length = 1 },
                    _ => null
                };
                if (leaf == null) continue;
                leaf.Start = pos; leaf.Field = field;
                pos += leaf.Length;
                res.Add(leaf);
            }
        }
        return res;
    }

    public static int Length(Paragraph p) { var l = Leaves(p); return l.Count == 0 ? 0 : l[^1].End; }

    public static string Text(Paragraph p) => string.Concat(Leaves(p).Select(l => l.Kind switch { LeafKind.Text => l.Text, LeafKind.Tab => "\t", LeafKind.Image => "￼", _ => "\n" }));
}
