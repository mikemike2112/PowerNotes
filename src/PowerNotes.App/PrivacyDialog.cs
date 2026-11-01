using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using Microsoft.Win32;
using PowerNotes.Core.Privacy;

namespace PowerNotes.App;

/// <summary>Scudo privacy: mostra cosa trova nel documento aperto (valori reali, una casella per voce) e crea una COPIA ripulita. Il file aperto e l'originale non vengono mai modificati.</summary>
static class PrivacyDialog
{
    public static void Show(Window owner, byte[] currentBytes, string? currentPath, string displayName)
    {
        ScanResult scan;
        try { scan = PrivacyShield.Scan(currentBytes); }
        catch (Exception e) when (e is InvalidDataException or IOException or System.Xml.XmlException) { ThemedDialog.Show(owner, "Impossibile analizzare il documento: " + e.Message, "Scudo privacy", MessageBoxButton.OK, MessageBoxImage.Error); return; }

        var root = new StackPanel { MinWidth = 520 };
        TextBlock Muted(string t, double size = 12.5, Thickness? m = null) { var x = new TextBlock { Text = t, FontSize = size, TextWrapping = TextWrapping.Wrap, Margin = m ?? new Thickness(0, 6, 0, 0) }; x.SetResourceReference(TextBlock.ForegroundProperty, "MutedBrush"); return x; }
        var boxes = new List<(CheckBox Box, Finding F)>();
        var list = new StackPanel();
        if (scan.Findings.Count == 0) list.Children.Add(Muted("Niente trovato tra le voci che lo scudo sa riconoscere (autore, società, proprietà, autori di commenti e revisioni, percorsi locali, miniatura).", 13));
        foreach (var f in scan.Findings)
        {
            var cb = new CheckBox { IsChecked = true, Margin = new Thickness(0, 4, 0, 4), Content = new TextBlock { TextWrapping = TextWrapping.Wrap, Text = $"{f.Label}: {f.Value}" } };
            AutomationProperties.SetName(cb, $"{f.Label}: {f.Value}");
            boxes.Add((cb, f)); list.Children.Add(cb);
        }
        root.Children.Add(new ScrollViewer { Content = list, MaxHeight = 220, VerticalScrollBarVisibility = ScrollBarVisibility.Auto });

        RadioButton? anon = null, remove = null;
        if (scan.Findings.Any(f => f.Kind == PrivacyKind.CommentAuthor))
        {
            root.Children.Add(Muted("Commenti:", 12.5, new Thickness(0, 12, 0, 2)));
            anon = new RadioButton { Content = "Anonimizza l'autore (predefinito, il contenuto non cambia)", IsChecked = true, Margin = new Thickness(0, 2, 0, 2), GroupName = "cm" };
            remove = new RadioButton { Content = "Rimuovi i commenti", Margin = new Thickness(0, 2, 0, 2), GroupName = "cm" };
            AutomationProperties.SetName(anon, "Anonimizza l'autore dei commenti"); AutomationProperties.SetName(remove, "Rimuovi i commenti");
            root.Children.Add(anon); root.Children.Add(remove);
        }

        root.Children.Add(Muted("Avvisi: cosa lo scudo NON pulisce", 12.5, new Thickness(0, 12, 0, 2)));
        foreach (var w in scan.Warnings) root.Children.Add(Muted("• " + w, 12, new Thickness(0, 2, 0, 0)));

        var result = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 12, 0, 0), FontSize = 12.5 };
        result.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush"); AutomationProperties.SetName(result, "Esito");
        var make = new Button { Content = "Crea copia pulita…", Padding = new Thickness(14, 6, 14, 6), Margin = new Thickness(0, 14, 0, 0), HorizontalAlignment = HorizontalAlignment.Left, IsEnabled = scan.Findings.Count > 0 };
        AutomationProperties.SetName(make, "Crea copia pulita"); make.SetResourceReference(Control.BorderBrushProperty, "AccentBrush"); make.BorderThickness = new Thickness(1);
        make.Click += (_, _) =>
        {
            var ids = boxes.Where(b => b.Box.IsChecked == true).Select(b => b.F.Id).ToHashSet();
            if (ids.Count == 0) { result.Text = "Seleziona almeno una voce."; return; }
            string folder = currentPath != null ? Path.GetDirectoryName(currentPath)! : PowerNotes.Core.Edit.SaveFlow.SuggestedFolder();
            var dlg = new SaveFileDialog { Filter = "Documenti (*.docx)|*.docx", DefaultExt = ".docx", FileName = Path.GetFileNameWithoutExtension(displayName) + " (pulito).docx", OverwritePrompt = true };
            if (Directory.Exists(folder)) dlg.InitialDirectory = folder;
            if (dlg.ShowDialog(owner) != true) return;
            if (currentPath != null && string.Equals(Path.GetFullPath(dlg.FileName), Path.GetFullPath(currentPath), StringComparison.OrdinalIgnoreCase))
            { result.Text = "Scegli un nome diverso: lo scudo non modifica mai il file aperto."; return; }
            try
            {
                var res = PrivacyShield.Clean(currentBytes, ids, remove?.IsChecked == true ? CommentMode.Remove : CommentMode.Anonymize);
                var tmp = dlg.FileName + ".tmp"; File.WriteAllBytes(tmp, res.Copy); File.Move(tmp, dlg.FileName, true);
                var sb = new System.Text.StringBuilder($"Copia creata: {dlg.FileName}{Environment.NewLine}Fatto: {string.Join("; ", res.Done)}.");
                sb.Append(res.After.Findings.Count == 0 ? $"{Environment.NewLine}Ricontrollo della copia: nessuna delle voci riconosciute è rimasta." : $"{Environment.NewLine}Ricontrollo della copia, ancora presente: {string.Join("; ", res.After.Findings.Select(f => f.Label + " «" + f.Value + "»"))}.");
                sb.Append($"{Environment.NewLine}Restano valide le riserve dell'elenco «Avvisi».");
                result.Text = sb.ToString();
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException or InvalidDataException) { result.Text = "Copia non creata: " + e.Message; }
        };
        root.Children.Add(make); root.Children.Add(result);
        ThemedDialog.ShowCustom(owner, "Scudo privacy", "Analisi del documento aperto, tutto in locale. Il file aperto non viene toccato: si crea una copia.", root);
    }
}
