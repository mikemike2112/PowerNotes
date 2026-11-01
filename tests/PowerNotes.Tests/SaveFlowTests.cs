using PowerNotes.Core.Edit;

namespace PowerNotes.Tests;

public class SaveFlowTests
{
    static readonly string Dir = Path.Combine(Path.GetTempPath(), "pn-save");

    [Theory]
    [InlineData("", false)] [InlineData("  ", false)] [InlineData("a/b", false)] [InlineData("a:b", false)] [InlineData("CON", false)]
    [InlineData("nome.", false)] [InlineData("Relazione", true)] [InlineData("a b.docx", true)]
    public void ValidateName(string n, bool ok) => Assert.Equal(ok, SaveFlow.ValidateName(n) == null);

    [Fact] public void Resolve_AddsExtension() => Assert.EndsWith("x.docx", SaveFlow.Resolve("x", Dir).Path);

    [Fact]
    public void Cancel_DoesNotWriteNorProceed()
    {
        int w = 0;
        Assert.False(SaveFlow.Run(_ => new(SaveChoice.Cancel), _ => true, _ => w++, null));
        Assert.Equal(0, w);
    }

    [Fact]
    public void DontSave_ProceedsWithoutWriting()
    {
        int w = 0;
        Assert.True(SaveFlow.Run(_ => new(SaveChoice.DontSave), _ => true, _ => w++, null));
        Assert.Equal(0, w);
    }

    [Fact]
    public void Save_InvalidName_ReportsErrorThenRetries()
    {
        var errors = new List<string?>(); var saved = new List<string>(); int n = 0;
        var ok = SaveFlow.Run(e => { errors.Add(e); return n++ == 0 ? new(SaveChoice.Save, "a/b", Dir) : new(SaveChoice.Save, "ok", Dir); }, _ => true, saved.Add, null, _ => false);
        Assert.True(ok); Assert.Single(saved); Assert.Null(errors[0]); Assert.NotNull(errors[1]);
    }

    [Fact]
    public void Save_ExistingFile_AsksAndNeverOverwritesSilently()
    {
        var saved = new List<string>(); int asks = 0, n = 0;
        // prima l'utente rifiuta la sovrascrittura, poi annulla
        var ok = SaveFlow.Run(_ => n++ == 0 ? new(SaveChoice.Save, "x", Dir) : new(SaveChoice.Cancel), _ => { asks++; return false; }, saved.Add, null, _ => true);
        Assert.False(ok); Assert.Equal(1, asks); Assert.Empty(saved);
        // conferma: scrive
        Assert.True(SaveFlow.Run(_ => new(SaveChoice.Save, "x", Dir), _ => true, saved.Add, null, _ => true));
        Assert.Single(saved);
    }

    [Fact]
    public void Save_ToCurrentPath_NoOverwriteQuestion_AndWriteFailureIsReported()
    {
        var cur = Path.Combine(Dir, "x.docx"); int asks = 0, n = 0; string? last = null;
        Assert.True(SaveFlow.Run(_ => new(SaveChoice.Save, "x", Dir), _ => { asks++; return true; }, _ => { }, cur, _ => true));
        Assert.Equal(0, asks);
        var ok = SaveFlow.Run(e => { last = e; return n++ == 0 ? new(SaveChoice.Save, "y", Dir) : new(SaveChoice.Cancel); }, _ => true, _ => throw new IOException("bloccato"), null, _ => false);
        Assert.False(ok); Assert.Contains("bloccato", last);
    }
}
