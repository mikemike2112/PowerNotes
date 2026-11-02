using System.Diagnostics;

namespace PowerNotes.Tests;

/// <summary>Avvia la vera applicazione WPF in modalità --selftest: eventi reali su superficie e ribbon, salvataggio, PDF, XPS di stampa.</summary>
[Collection("OfficeCom"), Trait("Category", "UI")]
public class UiSelfTest
{
    [Fact]
    public void RealApp_SelfTest_Passes()
    {
        using var lk = UiLock.Acquire();
        var exe = Directory.GetFiles(Path.Combine(TestPaths.Root, "src", "PowerNotes.App", "bin"), "PowerNotes.exe", SearchOption.AllDirectories).OrderByDescending(File.GetLastWriteTimeUtc).First();
        var dir = Path.Combine(TestPaths.Root, "artifacts", "selftest");
        if (Directory.Exists(dir)) Directory.Delete(dir, true);
        using var p = Process.Start(new ProcessStartInfo(exe, $"--selftest \"{dir}\"") { UseShellExecute = false })!;
        if (!p.WaitForExit(120_000)) { p.Kill(true); Assert.Fail("autotest UI oltre il tempo limite"); }
        var log = File.Exists(Path.Combine(dir, "selftest.txt")) ? File.ReadAllText(Path.Combine(dir, "selftest.txt")) : File.ReadAllText(Path.Combine(dir, "selftest-errore.txt"));
        Assert.True(p.ExitCode == 0, log);
        Assert.DoesNotContain("FALLITO", log);
    }
}
