using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Windows.Shell;
using PowerNotes.Core.Edit;

namespace PowerNotes.App;

/// <summary>Finestra di dialogo propria: senza bordo di sistema, angoli arrotondati, colori del tema attivo, accento dell'app.</summary>
public sealed class ThemedDialog : Window
{
    [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int val, int size);
    readonly StackPanel _body = new();
    readonly StackPanel _buttons = new() { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 22, 0, 0) };
    Button? _primary, _cancel;
    Control? _focusFirst;
    public MessageBoxResult Result { get; private set; } = MessageBoxResult.Cancel;

    ThemedDialog(Window? owner, string title, string message, string icon)
    {
        Owner = owner is { IsVisible: true } ? owner : null;
        WindowStartupLocation = Owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
        // Niente finestra trasparente (layered) né effetti software: il testo resta nitido (ClearType) anche con scala Windows > 100%.
        // Ombra e angoli arrotondati li fornisce DWM (Windows 11); su Windows 10 gli angoli restano squadrati.
        WindowStyle = WindowStyle.None; ShowInTaskbar = false; ResizeMode = ResizeMode.NoResize; UseLayoutRounding = true; SnapsToDevicePixels = true;
        WindowChrome.SetWindowChrome(this, new WindowChrome { CaptionHeight = 0, ResizeBorderThickness = new Thickness(0), GlassFrameThickness = new Thickness(1), UseAeroCaptionButtons = false, CornerRadius = new CornerRadius(0) });
        SetResourceReference(BackgroundProperty, "PanelBrush");
        SourceInitialized += (_, _) => { int pref = 2; DwmSetWindowAttribute(new WindowInteropHelper(this).Handle, 33, ref pref, 4); };
        SizeToContent = SizeToContent.WidthAndHeight; MinWidth = 380; MaxWidth = 560; Title = title;
        AutomationProperties.SetName(this, title);

        var titleText = new TextBlock { Text = title, FontSize = 17, FontWeight = FontWeights.SemiBold, TextWrapping = TextWrapping.Wrap };
        titleText.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
        var head = new StackPanel { Children = { titleText } };
        if (message.Length > 0)
        {
            var msg = new TextBlock { Text = message, FontSize = 13.5, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 8, 0, 0) };
            msg.SetResourceReference(TextBlock.ForegroundProperty, "MutedBrush");
            head.Children.Add(msg);
        }
        var ic = Icons.Make(icon, 30, icon == "dlg-error" ? "PowerBrush" : "AccentBrush"); ic.VerticalAlignment = VerticalAlignment.Top; ic.Margin = new Thickness(0, 2, 16, 0);
        var top = new DockPanel(); DockPanel.SetDock(ic, Dock.Left); top.Children.Add(ic); top.Children.Add(head);
        var root = new StackPanel { Margin = new Thickness(26, 24, 26, 20) };
        _body.Margin = new Thickness(0, 12, 0, 0);
        root.Children.Add(top); root.Children.Add(_body); root.Children.Add(_buttons);
        var card = new Border { BorderThickness = new Thickness(1), Child = root };
        card.SetResourceReference(Border.BackgroundProperty, "PanelBrush"); card.SetResourceReference(Border.BorderBrushProperty, "LineBrush");
        Content = card;
        MouseLeftButtonDown += (_, e) => { if (e.OriginalSource is not TextBox) { try { DragMove(); } catch (InvalidOperationException) { } } };
        KeyDown += (_, e) => { if (e.Key == Key.Escape && _cancel != null) { _cancel.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); e.Handled = true; } };
        Loaded += (_, _) => (_focusFirst ?? _primary)?.Focus();
    }

    Button AddButton(string text, MessageBoxResult r, bool primary, bool cancel, Action? beforeClose = null)
    {
        var b = new Button { Content = text, MinWidth = 92, Padding = new Thickness(14, 6, 14, 6), Margin = new Thickness(8, 0, 0, 0), IsDefault = primary };
        AutomationProperties.SetName(b, text);
        if (primary) { b.SetResourceReference(Control.BorderBrushProperty, "AccentBrush"); b.SetResourceReference(Control.BackgroundProperty, "AccentBrush"); b.Foreground = Brushes.Black; b.FontWeight = FontWeights.SemiBold; _primary = b; }
        else b.SetResourceReference(Control.BorderBrushProperty, "LineBrush");
        if (cancel) _cancel = b;
        b.Click += (_, _) => { beforeClose?.Invoke(); Result = r; Close(); };
        _buttons.Children.Add(b);
        return b;
    }

    static string IconKey(MessageBoxImage i) => i switch { MessageBoxImage.Error => "dlg-error", MessageBoxImage.Warning => "dlg-warn", MessageBoxImage.Question => "dlg-question", _ => "dlg-info" };

    /// <summary>Sostituto di MessageBox.Show.</summary>
    public static MessageBoxResult Show(Window? owner, string message, string title, MessageBoxButton buttons = MessageBoxButton.OK, MessageBoxImage image = MessageBoxImage.Information)
    {
        var d = new ThemedDialog(owner, title, message, IconKey(image));
        switch (buttons)
        {
            case MessageBoxButton.OK: d.AddButton("OK", MessageBoxResult.OK, true, true); break;
            case MessageBoxButton.OKCancel: d.AddButton("Annulla", MessageBoxResult.Cancel, false, true); d.AddButton("OK", MessageBoxResult.OK, true, false); break;
            case MessageBoxButton.YesNo: d.AddButton("No", MessageBoxResult.No, false, true); d.AddButton("Sì", MessageBoxResult.Yes, true, false); break;
            default: d.AddButton("Annulla", MessageBoxResult.Cancel, false, true); d.AddButton("No", MessageBoxResult.No, false, false); d.AddButton("Sì", MessageBoxResult.Yes, true, false); break;
        }
        d.ShowDialog();
        return d.Result;
    }

    public static MessageBoxResult Show(string message, string title, MessageBoxButton b, MessageBoxImage i) => Show(null, message, title, b, i);

    TextBox AddNameAndPath(string name, string folder, Func<(string Folder, string Name), string?>? browse)
    {
        TextBlock Lbl(string t) { var l = new TextBlock { Text = t, FontSize = 12, Margin = new Thickness(0, 14, 0, 4) }; l.SetResourceReference(TextBlock.ForegroundProperty, "MutedBrush"); return l; }
        var nameBox = new TextBox { Text = name, Height = 30, VerticalContentAlignment = VerticalAlignment.Center };
        AutomationProperties.SetName(nameBox, "Nome file");
        var pathBox = new TextBox { Text = folder, Height = 30, VerticalContentAlignment = VerticalAlignment.Center };
        AutomationProperties.SetName(pathBox, "Percorso");
        var browseBtn = new Button { Content = browse == null ? "Altre opzioni…" : "Sfoglia…", Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(12, 4, 12, 4) };
        AutomationProperties.SetName(browseBtn, "Sfoglia");
        browseBtn.Click += (_, _) =>
        {
            var r = browse?.Invoke((pathBox.Text, nameBox.Text));
            if (r != null) { pathBox.Text = System.IO.Path.GetDirectoryName(r) ?? pathBox.Text; nameBox.Text = System.IO.Path.GetFileNameWithoutExtension(r); }
        };
        var row = new DockPanel(); DockPanel.SetDock(browseBtn, Dock.Right); row.Children.Add(browseBtn); row.Children.Add(pathBox);
        _body.Children.Add(Lbl("Nome file")); _body.Children.Add(nameBox); _body.Children.Add(Lbl("Percorso")); _body.Children.Add(row);
        _focusFirst = nameBox;
        Loaded += (_, _) => nameBox.SelectAll();
        _fields = (nameBox, pathBox);
        return nameBox;
    }
    (TextBox Name, TextBox Path)? _fields;

    void AddError(string error)
    {
        var e = new TextBlock { Text = error, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 12, 0, 0), FontSize = 12.5 };
        e.SetResourceReference(TextBlock.ForegroundProperty, "PowerBrush"); AutomationProperties.SetName(e, "Errore"); _body.Children.Add(e);
    }

    /// <summary>
    /// Richiesta di salvataggio. askDiscard: "Vuoi salvare le modifiche a «nome»?" con Salva / Non salvare / Annulla (nome e percorso solo se il documento non ha un file);
    /// altrimenti pannello "Salva con nome" (Salva / Annulla), sempre con nome e percorso.
    /// </summary>
    public static SavePromptResult SavePrompt(Window? owner, string docName, bool askDiscard, bool needsPath, string folder, string? error, Func<(string Folder, string Name), string?> browse)
    {
        var d = askDiscard ? new ThemedDialog(owner, $"Vuoi salvare le modifiche a «{docName}»?", "Se non salvi, le modifiche andranno perse.", "dlg-question")
                           : new ThemedDialog(owner, "Salva con nome", "", "dlg-question");
        if (needsPath || !askDiscard) d.AddNameAndPath(System.IO.Path.GetFileNameWithoutExtension(docName), folder, browse);
        if (error != null) d.AddError(error);
        var choice = SaveChoice.Cancel; string name = "", fold = "";
        d.AddButton("Annulla", MessageBoxResult.Cancel, false, true, () => choice = SaveChoice.Cancel);
        if (askDiscard) d.AddButton("Non salvare", MessageBoxResult.No, false, false, () => choice = SaveChoice.DontSave);
        d.AddButton("Salva", MessageBoxResult.Yes, true, false, () => { choice = SaveChoice.Save; name = d._fields?.Name.Text ?? ""; fold = d._fields?.Path.Text ?? ""; });
        d.ShowDialog();
        return new SavePromptResult(choice, name, fold);
    }

    /// <summary>Conferma con pulsante primario personalizzato (es. «Applica») e Annulla.</summary>
    public static bool Confirm(Window? owner, string title, string message, string primary)
    {
        var d = new ThemedDialog(owner, title, message, "dlg-info");
        d.AddButton("Annulla", MessageBoxResult.Cancel, false, true); d.AddButton(primary, MessageBoxResult.OK, true, false);
        d.ShowDialog();
        return d.Result == MessageBoxResult.OK;
    }

    /// <summary>Finestra con contenuto proprio e un solo pulsante «Chiudi».</summary>
    public static void ShowCustom(Window? owner, string title, string intro, FrameworkElement content, double maxWidth = 680)
    {
        var d = new ThemedDialog(owner, title, intro, "dlg-info") { MaxWidth = maxWidth, MinWidth = Math.Min(maxWidth, 560) };
        d._body.Children.Add(content);
        d.AddButton("Chiudi", MessageBoxResult.OK, true, true);
        d.ShowDialog();
    }

    /// <summary>Mostra il pannello senza bloccare (solo per screenshot e automazione).</summary>
    public static ThemedDialog PreviewSave(Window owner, string docName, bool topmost = false)
    {
        var d = new ThemedDialog(owner, $"Vuoi salvare le modifiche a «{docName}»?", "Se non salvi, le modifiche andranno perse.", "dlg-question");
        d.AddNameAndPath(System.IO.Path.GetFileNameWithoutExtension(docName), SaveFlow.SuggestedFolder(), _ => null);
        d.AddButton("Annulla", MessageBoxResult.Cancel, false, true); d.AddButton("Non salvare", MessageBoxResult.No, false, false); d.AddButton("Salva", MessageBoxResult.Yes, true, false);
        d.Topmost = topmost; d.Show();
        return d;
    }
}
