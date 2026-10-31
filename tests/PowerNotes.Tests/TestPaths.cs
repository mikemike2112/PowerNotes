using System.Text.Json;

namespace PowerNotes.Tests;

static class TestPaths
{
    public static string Root { get; } = FindRoot();
    public static string Corpus => Path.Combine(Root, "tests", "corpus", "office-made");
    public static string Tools => Path.Combine(Root, "tools");

    static string FindRoot()
    {
        var d = new DirectoryInfo(AppContext.BaseDirectory);
        while (d != null && !File.Exists(Path.Combine(d.FullName, "PowerNotes.sln"))) d = d.Parent;
        return d?.FullName ?? throw new InvalidOperationException("Radice del repository non trovata.");
    }

    public static IEnumerable<string> CorpusFiles() => Directory.GetFiles(Corpus, "*.docx").OrderBy(x => x);

    public static IEnumerable<object[]> CorpusNames() => CorpusFiles().Select(f => new object[] { Path.GetFileNameWithoutExtension(f) });

    static JsonDocument? _inspect;
    public static JsonElement Expected(string name)
    {
        _inspect ??= JsonDocument.Parse(File.ReadAllText(Path.Combine(Root, "tests", "corpus", "office-made.inspect.json")));
        var path = Path.Combine(Corpus, name + ".docx");
        foreach (var p in _inspect.RootElement.EnumerateObject())
            if (string.Equals(p.Name, path, StringComparison.OrdinalIgnoreCase)) return p.Value;
        throw new KeyNotFoundException(name);
    }
}
