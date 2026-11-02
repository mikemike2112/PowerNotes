using System.Diagnostics;
using System.Text.Json;

namespace PowerNotes.Tests;

/// <summary>Input reale (mouse) sulla scheda Powers: tema, persistenza dopo riavvio, impostazioni corrotte. Screenshot in artifacts/powers.</summary>
[Collection("OfficeCom"), Trait("Category", "UI")]
public class PowersUiTest
{
    [Fact]
    public async Task Switch_ChangesTheme_Persists_AndCorruptSettingsDoNotCrash()
    {
        using var lk = UiLock.Acquire();
        var exe = Directory.GetFiles(Path.Combine(TestPaths.Root, "src", "PowerNotes.App", "bin"), "PowerNotes.exe", SearchOption.AllDirectories).OrderByDescending(File.GetLastWriteTimeUtc).First();
        var dir = Path.Combine(TestPaths.Root, "artifacts", "powers"); var json = Path.Combine(dir, "result.json");
        Directory.CreateDirectory(dir); if (File.Exists(json)) File.Delete(json);
        var psi = new ProcessStartInfo("powershell", $"-NoProfile -ExecutionPolicy Bypass -File \"{Path.Combine(TestPaths.Tools, "ui-powers-test.ps1")}\" -Exe \"{exe}\" -OutDir \"{dir}\" -OutJson \"{json}\"") { UseShellExecute = false, RedirectStandardError = true, CreateNoWindow = true };
        using var p = Process.Start(psi)!;
        var err = p.StandardError.ReadToEndAsync();
        if (!p.WaitForExit(240_000)) { p.Kill(true); Assert.Fail("test UI Powers oltre il tempo limite"); }
        Assert.True(File.Exists(json), "nessun risultato: " + await err);
        var r = JsonDocument.Parse(File.ReadAllText(json)).RootElement;
        Assert.Equal("On", r.GetProperty("firstRunState").GetString());        // predefinito scuro
        Assert.Equal("Off", r.GetProperty("afterClickState").GetString());     // clic reale → chiaro
        Assert.Contains("\"Dark\":false", r.GetProperty("savedJson").GetString());
        Assert.Equal("Off", r.GetProperty("restartState").GetString());        // persistenza dopo riavvio
        Assert.True(r.GetProperty("corruptAlive").GetBoolean());
        Assert.Equal("On", r.GetProperty("corruptState").GetString());         // corrotto → predefinito scuro
        foreach (var f in new[] { "powers-scuro.png", "powers-chiaro.png", "riavvio-chiaro.png" }) Assert.True(new FileInfo(Path.Combine(dir, f)).Length > 5000);
    }
}
