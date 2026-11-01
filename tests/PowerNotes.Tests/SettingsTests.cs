using PowerNotes.Core.Settings;

namespace PowerNotes.Tests;

public class SettingsTests
{
    [Fact]
    public void Default_IsDark_RoundTrips_AndCorruptFileIsHandled()
    {
        var dir = Path.Combine(Path.GetTempPath(), "pn-set-" + Guid.NewGuid().ToString("N")); var f = Path.Combine(dir, "settings.json");
        try
        {
            Assert.True(AppSettings.Load(f).Dark);                       // assente → scuro
            new AppSettings { Dark = false }.Save(f);
            Assert.False(AppSettings.Load(f).Dark);
            File.WriteAllText(f, "{ non json ::: ");
            Assert.True(AppSettings.Load(f).Dark);                       // corrotto → predefinito, nessuna eccezione
            File.WriteAllBytes(f, [0xFF, 0x00, 0x13]);
            Assert.True(AppSettings.Load(f).Dark);
            File.WriteAllText(f, "null");
            Assert.True(AppSettings.Load(f).Dark);
        }
        finally { if (Directory.Exists(dir)) Directory.Delete(dir, true); }
    }
}
