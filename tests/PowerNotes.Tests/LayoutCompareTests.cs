using PowerNotes.Core.Model;
using PowerNotes.Core.Edit;
using Xunit.Abstractions;

namespace PowerNotes.Tests;

public class LayoutCompareTests(ITestOutputHelper o)
{
    [Fact]
    public void Dump()
    {
        var only = Environment.GetEnvironmentVariable("PN_DUMP");
        int ok = 0, tot = 0;
        foreach (var f in TestPaths.CorpusFiles())
        {
            var name = Path.GetFileNameWithoutExtension(f);
            if (only != null && !only.Split(',').Contains(name)) continue;
            using var d = WordDoc.Open(f);
            var res = d.Layout();
            var exp = TestPaths.Expected(name);
            int wp = exp.GetProperty("pages").GetInt32();
            tot++; if (wp == res.Pages.Count) ok++;
            o.WriteLine($"{name,-28} word={wp,3} mine={res.Pages.Count,3} {(wp == res.Pages.Count ? "" : "<<<")}");
            if (only != null)
            {
                var ps = d.Paragraphs();
                var lines = exp.GetProperty("lines");
                for (int i = 0; i < ps.Count && i < lines.GetArrayLength(); i++)
                {
                    var l = res.LinesOf(ps[i]).FirstOrDefault();
                    var wl = lines[i];
                    o.WriteLine($"  p{i,3} word pg{wl[0].GetInt32()} y={wl[1].GetDouble(),7:F2} x={wl[2].GetDouble(),6:F1} | mine pg{(l?.PageIndex ?? -1) + 1} y={l?.Y ?? -1,7:F2} x={l?.Left ?? -1,6:F1}  '{ParaText.Text(ps[i])[..Math.Min(20, ParaText.Text(ps[i]).Length)]}'");
                }
            }
        }
        o.WriteLine($"MATCH {ok}/{tot}");
    }
}
