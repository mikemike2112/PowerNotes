namespace PowerNotes.Core.Edit;

public enum SaveChoice { Save, DontSave, Cancel }

/// <summary>Esito della richiesta di salvataggio: scelta, nome file e cartella (usati solo con Save).</summary>
public sealed record SavePromptResult(SaveChoice Choice, string Name = "", string Folder = "");

/// <summary>Logica della richiesta di salvataggio, senza interfaccia: validazione del nome, sovrascrittura, esiti. Nulla viene scritto finché non c'è Save.</summary>
public static class SaveFlow
{
    static readonly string[] Reserved = ["CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"];

    /// <summary>Messaggio d'errore se il nome non è valido, altrimenti null.</summary>
    public static string? ValidateName(string? name)
    {
        name = name?.Trim() ?? "";
        if (name.Length == 0) return "Inserisci un nome per il file.";
        if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return "Il nome non può contenere i caratteri: barre, due punti, * ? virgolette, < > |";
        if (name.TrimEnd('.', ' ').Length != name.Length) return "Il nome non può terminare con un punto o uno spazio.";
        var stem = Path.GetFileNameWithoutExtension(name).ToUpperInvariant();
        if (Reserved.Contains(stem)) return "Questo nome è riservato dal sistema.";
        if (name.Length > 200) return "Il nome è troppo lungo.";
        return null;
    }

    /// <summary>Percorso completo; aggiunge .docx se manca. Errore (con path null) se nome o cartella non sono validi.</summary>
    public static (string? Path, string? Error) Resolve(string? name, string? folder)
    {
        var e = ValidateName(name);
        if (e != null) return (null, e);
        if (string.IsNullOrWhiteSpace(folder)) return (null, "Scegli una cartella.");
        name = name!.Trim();
        if (!name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) name += ".docx";
        try { return (Path.GetFullPath(Path.Combine(folder.Trim(), name)), null); }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException) { return (null, "Cartella non valida."); }
    }

    /// <summary>Cartella proposta per un documento senza file: Documenti dell'utente. È solo una proposta, non viene creata né scritta.</summary>
    public static string SuggestedFolder() => Environment.GetEnvironmentVariable("POWERNOTES_DEFAULT_FOLDER") is { Length: > 0 } f ? f : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);   // la variabile serve solo per gli screenshot

    /// <summary>
    /// Esegue il flusso. prompt: mostra la richiesta (ripetuta finché Salva non va a buon fine o l'utente sceglie altro).
    /// confirmOverwrite: true per sovrascrivere un file esistente diverso dal documento. write: salva atomicamente.
    /// Restituisce true se si può procedere (salvato o Non salvare), false se annullato.
    /// </summary>
    public static bool Run(Func<string?, SavePromptResult> prompt, Func<string, bool> confirmOverwrite, Action<string> write, string? currentPath, Func<string, bool>? exists = null)
    {
        exists ??= File.Exists;
        string? error = null;
        while (true)
        {
            var r = prompt(error);
            switch (r.Choice)
            {
                case SaveChoice.Cancel: return false;
                case SaveChoice.DontSave: return true;
            }
            string path;
            if (currentPath != null && string.IsNullOrEmpty(r.Name)) path = currentPath;
            else
            {
                var (p, err) = Resolve(r.Name, r.Folder);
                if (p == null) { error = err; continue; }
                path = p;
                bool same = currentPath != null && string.Equals(Path.GetFullPath(currentPath), path, StringComparison.OrdinalIgnoreCase);
                if (!same && exists(path) && !confirmOverwrite(path)) { error = null; continue; }
            }
            try { write(path); return true; }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { error = "Salvataggio non riuscito, nulla è stato modificato: " + ex.Message; }
        }
    }
}
