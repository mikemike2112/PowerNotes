using System.Text.Json;

namespace PowerNotes.Core.Settings;

/// <summary>Impostazioni utente in %APPDATA%\PowerNotes\settings.json. File assente o corrotto → valori predefiniti, mai un crash.</summary>
public sealed class AppSettings
{
    public bool Dark { get; set; } = true;   // predefinito: scuro alla prima apertura

    public static string DefaultPath => Environment.GetEnvironmentVariable("POWERNOTES_SETTINGS")
        ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PowerNotes", "settings.json");

    public static AppSettings Load(string? path = null)
    {
        try
        {
            var json = File.ReadAllText(path ?? DefaultPath);
            return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or JsonException or NotSupportedException or ArgumentException)
        {
            return new AppSettings();
        }
    }

    public void Save(string? path = null)
    {
        path ??= DefaultPath;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            var tmp = path + ".tmp";
            File.WriteAllText(tmp, JsonSerializer.Serialize(this));
            File.Move(tmp, path, true);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { /* ponytail: impostazioni non salvate, l'app continua */ }
    }
}
