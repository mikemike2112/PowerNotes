using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Shell;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Model;
using PowerNotes.Core.Settings;

namespace PowerNotes.App;

/// <summary>Costruzione dell'interfaccia: barra del titolo, schede, ribbon, righelli, stato, scheda File e scheda Powers.</summary>
public partial class MainWindow
{
    ToggleButton BoldBtn = null!, ItalicBtn = null!, UnderBtn = null!, StrikeBtn = null!, AlLeft = null!, AlCenter = null!, AlRight = null!, AlJust = null!;
    ComboBox FontBox = null!, SizeBox = null!, LineBox = null!, StyleBox = null!;
    ToggleButton? DarkSwitch;
    TextBlock TitleLabel = null!;
    Slider _zoomSlider = null!;
    TextBlock _zoomPct = null!;
    TextBox _indL = null!, _indR = null!, _spB = null!, _spA = null!;
    readonly List<(string Name, RibbonBar Content, bool Contextual, TextBlock Header, Border Frame)> _tabs = new();
    List<UIElement>? _curLabels;
    Button? _maxBtn;
    readonly List<(string Label, Action Act, Button Btn)> _cmds = new();
    int _tab;
    Button? _printViewBtn;
    const string Soon = "Non ancora disponibile";

    // ---------------------------------------------------------------- costruzione
    void BuildUi()
    {
        WindowStyle = WindowStyle.None;
        WindowChrome.SetWindowChrome(this, new WindowChrome { CaptionHeight = 40, ResizeBorderThickness = new Thickness(6), GlassFrameThickness = new Thickness(0), UseAeroCaptionButtons = false, CornerRadius = new CornerRadius(0) });
        StateChanged += (_, _) => ((FrameworkElement)Content).Margin = WindowState == WindowState.Maximized ? new Thickness(7) : new Thickness(0);
        View.Scroller = Scroller;
        View.LinkClicked += OpenLink;

        BuildTitleBar();
        BuildRibbon();
        BuildStatusBar();
        BuildFilePage();
        CollapseBtn.Content = Icons.Make("caret", 14);
        CollapseBtn.Click += (_, _) => { bool c = RibbonFrame.Height > 40; RibbonFrame.Height = c ? 8 : 126; RibbonHost.Visibility = c ? Visibility.Collapsed : Visibility.Visible; };
        Scroller.ScrollChanged += (_, _) => UpdateRulers();
        SizeChanged += (_, _) => { UpdateRulers(); Adapt(); };
        View.StateChanged += UpdateRulers;
        SelectTab(0);
    }

    static Border Hover(UIElement child) => new() { Child = child };

    Button Make(string? label, string? icon, double iconSize, bool large, Action act, string? tip, bool enabled)
    {
        var sp = new StackPanel { Orientation = large ? Orientation.Vertical : Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center };
        if (icon != null) { var ic = Icons.Make(icon, iconSize); if (!large && label != null) ic.Margin = new Thickness(0, 0, 6, 0); if (large) ic.Margin = new Thickness(0, 2, 0, 3); sp.Children.Add(ic); }
        if (label != null)
        {
            var tl = new TextBlock { Text = label, FontSize = 12, TextAlignment = TextAlignment.Center, TextWrapping = TextWrapping.Wrap, MaxWidth = large ? 76 : 200, VerticalAlignment = VerticalAlignment.Center };
            sp.Children.Add(tl);
            if (icon != null) _curLabels?.Add(tl);   // nascosta quando il gruppo si compatta
        }
        var b = new Button { Content = sp, Padding = large ? new Thickness(6, 2, 6, 2) : new Thickness(5, 3, 5, 3), MinWidth = large ? 54 : 24, HorizontalContentAlignment = large ? HorizontalAlignment.Center : HorizontalAlignment.Left };
        System.Windows.Automation.AutomationProperties.SetName(b, label ?? tip ?? icon ?? "");
        b.ToolTip = enabled ? (tip ?? label) : Soon;
        b.IsEnabled = enabled;
        if (enabled) { b.Click += (_, _) => act(); _cmds.Add((label ?? tip ?? "", act, b)); }
        return b;
    }

    Button Large(Panel host, string label, string icon, Action act, string? tip = null, bool enabled = true) { var b = Make(label, icon, 30, true, act, tip, enabled); host.Children.Add(b); return b; }
    Button Small(Panel host, string? label, string icon, Action act, string? tip = null, bool enabled = true) { var b = Make(label, icon, 16, false, act, tip ?? label, enabled); host.Children.Add(b); return b; }
    Button Stub(Panel host, string label, string icon, bool large = false) => large ? Large(host, label, icon, () => { }, null, false) : Small(host, label, icon, () => { }, null, false);

    ToggleButton Toggle(Panel host, string text, string tip, Action act, string? icon = null, FontWeight? w = null, FontStyle? st = null, TextDecorationCollection? deco = null)
    {
        var tb = new TextBlock { Text = text, FontSize = 13, FontWeight = w ?? FontWeights.Normal, FontStyle = st ?? FontStyles.Normal, TextDecorations = deco };
        object content = icon != null ? Icons.Make(icon, 16) : tb;
        var b = new ToggleButton { Content = content, ToolTip = tip, Padding = new Thickness(5, 3, 5, 3), MinWidth = 26 };
        System.Windows.Automation.AutomationProperties.SetName(b, tip);
        b.Click += (_, _) => act();
        host.Children.Add(b); _cmds.Add((tip, act, new Button()));
        return b;
    }

    static StackPanel Row(params UIElement[] els) { var p = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 1, 0, 1) }; foreach (var e in els) p.Children.Add(e); return p; }
    static StackPanel Col(params UIElement[] els) { var p = new StackPanel { Orientation = Orientation.Vertical, VerticalAlignment = VerticalAlignment.Top }; foreach (var e in els) p.Children.Add(e); return p; }

    StackPanel Group(RibbonBar tab, string name, Action? launcher = null)
    {
        var content = new StackPanel { Orientation = Orientation.Horizontal };
        var g = new Grid { Margin = new Thickness(0, 0, 0, 0) };
        g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); g.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        g.Children.Add(content);
        var label = new TextBlock { Text = name, FontSize = 11, HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(6, 0, 6, 2) };
        label.SetResourceReference(TextBlock.ForegroundProperty, "MutedBrush"); Grid.SetRow(label, 1); g.Children.Add(label);
        if (launcher != null)
        {
            var l = new Button { Content = new TextBlock { Text = "⌟", FontSize = 10 }, MinWidth = 14, Padding = new Thickness(1), ToolTip = name + "…", HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Bottom };
            l.Click += (_, _) => launcher(); Grid.SetRow(l, 1); g.Children.Add(l);
        }
        var sep = new Border { Width = 1, Margin = new Thickness(4, 4, 0, 6) }; sep.SetResourceReference(Border.BackgroundProperty, "LineBrush");
        var info = tab.AddGroup(new StackPanel { Orientation = Orientation.Horizontal, Children = { g, sep } });
        _curLabels = info.Labels;
        return content;
    }

    RibbonBar NewTab(string name, bool contextual = false, bool power = false)
    {
        var content = new RibbonBar();
        var header = new TextBlock { Text = name, FontSize = 14, Padding = new Thickness(11, 6, 11, 4) };
        if (power) header.SetResourceReference(TextBlock.ForegroundProperty, "PowerBrush");
        else if (contextual) header.Foreground = new SolidColorBrush(Color.FromRgb(0x0a, 0x7f, 0x94));
        else header.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
        var frame = new Border { Child = header, BorderThickness = new Thickness(0, 0, 0, 2), BorderBrush = Brushes.Transparent, Margin = new Thickness(1, 0, 1, 0), Cursor = Cursors.Hand };
        int idx = _tabs.Count;
        frame.MouseLeftButtonDown += (_, _) => { if (CurrentFilePage) HideFile(); SelectTab(idx); };
        _tabs.Add((name, content, contextual, header, frame));
        TabRow.Children.Add(frame);
        if (contextual) frame.Visibility = Visibility.Collapsed;
        return content;
    }

    public bool CurrentFilePage => FilePage.Visibility == Visibility.Visible;

    public void SelectTab(int i)
    {
        if (i < 0 || i >= _tabs.Count) return;
        _tab = i;
        for (int k = 0; k < _tabs.Count; k++)
        {
            var f = _tabs[k].Frame; bool sel = k == i;
            if (sel) { if (_tabs[k].Name == "Powers") f.BorderBrush = (Brush)FindResource("PowerBrush"); else f.SetResourceReference(Border.BorderBrushProperty, "TextBrush"); }
            else f.BorderBrush = Brushes.Transparent;
            _tabs[k].Header.FontWeight = sel ? FontWeights.SemiBold : FontWeights.Normal;
        }
        RibbonHost.Content = _tabs[i].Content;
        _tabs[i].Content.Relayout();
        // la cornice del ribbon segue il colore della scheda Powers
        if (_tabs[i].Name == "Powers") RibbonFrame.BorderBrush = (Brush)FindResource("PowerBrush"); else RibbonFrame.SetResourceReference(Border.BorderBrushProperty, "LineBrush");
    }

    void SetTableTab(bool on)
    {
        int idx = _tabs.FindIndex(t => t.Contextual);
        if (idx < 0) return;
        _tabs[idx].Frame.Visibility = on ? Visibility.Visible : Visibility.Collapsed;
        if (!on && _tab == idx) SelectTab(0);
    }

    // ---------------------------------------------------------------- titolo
    void BuildTitleBar()
    {
        var icon = new Image { Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("pack://application:,,,/Assets/powernotes.ico")), Width = 20, Height = 20, Margin = new Thickness(0, 0, 14, 0) };
        QuickBar.Children.Add(icon);
        WindowChrome.SetIsHitTestVisibleInChrome(QuickBar, true);
        var save = Small(QuickBar, null, "save", () => SaveCore(false), "Salva (Ctrl+S)");
        var undo = Small(QuickBar, null, "undo", () => Do(() => _s.Undo()), "Annulla (Ctrl+Z)");
        var redo = Small(QuickBar, null, "redo", () => Do(() => _s.Redo()), "Ripeti (Ctrl+Y)");
        Small(QuickBar, null, "caret", () => { }, null, false);
        System.Windows.Automation.AutomationProperties.SetName(undo, "Annulla"); System.Windows.Automation.AutomationProperties.SetName(redo, "Ripeti"); System.Windows.Automation.AutomationProperties.SetName(save, "Salva");
        TitleLabel = new TextBlock { Margin = new Thickness(18, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center, FontSize = 13 };
        QuickBar.Children.Add(TitleLabel);

        SearchIcon.Content = Icons.Make("search", 16, "MutedBrush");
        WindowChrome.SetIsHitTestVisibleInChrome(SearchBox, true);
        SearchBox.TextChanged += (_, _) => SearchHint.Visibility = SearchBox.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
        SearchBox.KeyDown += (_, e) => { if (e.Key == Key.Enter) RunSearch(SearchBox.Text.Trim()); };

        Button Cap(string icon, Action a, bool close = false)
        {
            var b = new Button { Content = Icons.Make(icon, 14), Width = 46, Height = 40, Padding = new Thickness(0), Margin = new Thickness(0) };
            b.Click += (_, _) => a(); WindowChrome.SetIsHitTestVisibleInChrome(b, true); CaptionButtons.Children.Add(b); return b;
        }
        Cap("min", () => WindowState = WindowState.Minimized);
        _maxBtn = Cap("max", () => WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized);
        StateChanged += (_, _) => UpdateMaxButton();
        UpdateMaxButton();
        Cap("close", Close, true);
    }

    /// <summary>Adatta barra del titolo e barra di stato alla larghezza: titolo troncato, viste e testi secondari nascosti, zoom compatto.</summary>
    void Adapt()
    {
        double w = ActualWidth;
        TitleLabel.TextTrimming = TextTrimming.CharacterEllipsis; TitleLabel.MaxWidth = Math.Max(48, w * 0.2);
        ViewBar.Visibility = w < 960 ? Visibility.Collapsed : Visibility.Visible;
        WordsText.Visibility = w < 800 ? Visibility.Collapsed : Visibility.Visible;
        _zoomSlider.Width = w < 900 ? 56 : 120;
        QuickBar.Margin = new Thickness(w < 800 ? 6 : 12, 0, 0, 0);
    }

    /// <summary>Icona e tooltip del pulsante ingrandisci: un quadrato a finestra, due quadrati a schermo intero. Si aggiorna a ogni cambio di stato.</summary>
    void UpdateMaxButton()
    {
        if (_maxBtn == null) return;
        bool max = WindowState == WindowState.Maximized;
        _maxBtn.Content = Icons.Make(max ? "restore" : "max", 14);
        _maxBtn.ToolTip = max ? "Ripristina giù" : "Ingrandisci";
        System.Windows.Automation.AutomationProperties.SetName(_maxBtn, max ? "Ripristina giù" : "Ingrandisci");
    }

    /// <summary>Casella Cerca: esegue il primo comando abilitato il cui nome contiene il testo; altrimenti cerca il testo nel documento.</summary>
    void RunSearch(string q)
    {
        if (q.Length == 0) return;
        var hit = _cmds.FirstOrDefault(c => c.Btn.IsEnabled && c.Label.Contains(q, StringComparison.OrdinalIgnoreCase));
        if (hit.Act != null) { Do(hit.Act); return; }
        if (!_s.FindNext(q, false)) ThemedDialog.Show(this, "Nessun comando né testo trovato per \"" + q + "\".", "Cerca");
        View.Refresh(true);
    }

    // ---------------------------------------------------------------- ribbon
    void BuildRibbon()
    {
        // ====== File (apre la pagina a tutta pagina)
        var fileHeader = new TextBlock { Text = "File", FontSize = 14, Padding = new Thickness(11, 6, 11, 4) };
        fileHeader.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
        var fileFrame = new Border { Child = fileHeader, BorderThickness = new Thickness(0, 0, 0, 2), BorderBrush = Brushes.Transparent, Cursor = Cursors.Hand, Margin = new Thickness(1, 0, 1, 0) };
        System.Windows.Automation.AutomationProperties.SetName(fileFrame, "File");
        fileFrame.MouseLeftButtonDown += (_, _) => ShowFile("Informazioni");
        TabRow.Children.Add(fileFrame);

        // ====== Home
        var home = NewTab("Home");
        var g = Group(home, "Appunti", () => { });
        Large(g, "Incolla", "paste", () => Do(Paste), "Incolla (Ctrl+V)");
        var ap = new StackPanel(); Small(ap, "Taglia", "cut", () => Do(Cut), "Taglia (Ctrl+X)"); Small(ap, "Copia", "copy", () => Do(Copy), "Copia (Ctrl+C)"); Small(ap, "Copia formato", "brush", () => { }, null, false); g.Children.Add(ap);

        g = Group(home, "Carattere", () => CharDialog());
        FontBox = new ComboBox { Width = 150, IsEditable = true, ToolTip = "Tipo di carattere", Name = "FontBox", ItemsSource = System.Windows.Media.Fonts.SystemFontFamilies.Select(f => f.Source).OrderBy(x => x).ToList() };
        FontBox.DropDownClosed += (_, _) => Font_Changed(null, EventArgs.Empty);
        FontBox.KeyDown += (_, e) => { if (e.Key == Key.Enter) Font_Changed(null, EventArgs.Empty); };
        SizeBox = new ComboBox { Width = 56, IsEditable = true, ToolTip = "Dimensione", Name = "SizeBox", ItemsSource = new[] { "8", "9", "10", "10,5", "11", "12", "14", "16", "18", "20", "24", "28", "36", "48", "72" } };
        SizeBox.DropDownClosed += (_, _) => Size_Changed(null, EventArgs.Empty);
        SizeBox.KeyDown += (_, e) => { if (e.Key == Key.Enter) Size_Changed(null, EventArgs.Empty); };
        var r1 = Row(FontBox, SizeBox);
        Small(r1, null, "generic", () => Do(() => _s.StepSize(true)), "Ingrandisci carattere (Ctrl+>)").Content = new TextBlock { Text = "A▲", FontSize = 11 };
        Small(r1, null, "generic", () => Do(() => _s.StepSize(false)), "Riduci carattere (Ctrl+<)").Content = new TextBlock { Text = "A▼", FontSize = 11 };
        Small(r1, null, "generic", () => Do(() => _s.ClearFormatting()), "Cancella formattazione").Content = new TextBlock { Text = "A✕", FontSize = 11 };
        Small(r1, null, "generic", () => { }, null, false).Content = new TextBlock { Text = "Aa", FontSize = 12 };
        var r2 = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 1, 0, 1) };
        BoldBtn = Toggle(r2, "G", "Grassetto", () => Do(() => _s.ToggleBold()), null, FontWeights.Bold);
        ItalicBtn = Toggle(r2, "C", "Corsivo", () => Do(() => _s.ToggleItalic()), null, null, FontStyles.Italic);
        UnderBtn = Toggle(r2, "S", "Sottolineato", () => Do(() => _s.ToggleUnderline()), null, null, null, TextDecorations.Underline);
        StrikeBtn = Toggle(r2, "ab", "Barrato", () => Do(() => _s.ToggleStrike()), null, null, null, TextDecorations.Strikethrough);
        Small(r2, null, "generic", () => Do(() => _s.SetVertAlign(_s.AllRuns(f => f.VertAlign < 0) ? 0 : -1)), "Pedice").Content = new TextBlock { Text = "x₂", FontSize = 13 };
        Small(r2, null, "generic", () => Do(() => _s.SetVertAlign(_s.AllRuns(f => f.VertAlign > 0) ? 0 : 1)), "Apice").Content = new TextBlock { Text = "x²", FontSize = 13 };
        Small(r2, null, "generic", () => HighlightMenu(), "Evidenziazione").Content = new TextBlock { Text = "▮▾", FontSize = 12, Foreground = new SolidColorBrush(Color.FromRgb(0xe8, 0xc8, 0)) };
        Small(r2, null, "generic", () => ColorMenu(), "Colore carattere").Content = new TextBlock { Text = "A▾", FontSize = 13, Foreground = new SolidColorBrush(Color.FromRgb(0xd0, 0x20, 0x20)), FontWeight = FontWeights.Bold };
        g.Children.Add(Col(r1, r2));

        g = Group(home, "Paragrafo", () => ParaDialog_Click(this, new RoutedEventArgs()));
        var p1 = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 1, 0, 1) };
        Small(p1, null, "bullets", () => Do(() => _s.ToggleList("bullet")), "Elenchi puntati");
        Small(p1, null, "numbers", () => Do(() => _s.ToggleList("decimal")), "Elenchi numerati");
        Small(p1, null, "multilevel", () => { }, null, false);
        Small(p1, null, "outdent", () => Do(() => _s.ChangeIndent(-36)), "Riduci rientro");
        Small(p1, null, "indent", () => Do(() => _s.ChangeIndent(36)), "Aumenta rientro");
        Small(p1, null, "sort", () => { }, null, false);
        Small(p1, null, "pilcrow", () => { }, null, false);
        var p2 = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 1, 0, 1) };
        AlLeft = Toggle(p2, "", "Allinea a sinistra", () => Do(() => _s.SetAlignment(Align.Left)), "alignL");
        AlCenter = Toggle(p2, "", "Allinea al centro", () => Do(() => _s.SetAlignment(Align.Center)), "alignC");
        AlRight = Toggle(p2, "", "Allinea a destra", () => Do(() => _s.SetAlignment(Align.Right)), "alignR");
        AlJust = Toggle(p2, "", "Giustifica", () => Do(() => _s.SetAlignment(Align.Justify)), "alignJ");
        LineBox = new ComboBox { Width = 64, ToolTip = "Interlinea e spaziatura paragrafo", Name = "LineBox", ItemsSource = new[] { "1,0", "1,08", "1,15", "1,5", "2,0", "2,5", "3,0" } };
        LineBox.DropDownClosed += (_, _) => Line_Changed(null, EventArgs.Empty);
        p2.Children.Add(LineBox);
        Small(p2, null, "shading", () => { }, null, false); Small(p2, null, "borders", () => { }, null, false);
        g.Children.Add(Col(p1, p2));

        g = Group(home, "Stili", () => { });
        var gal = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(2, 4, 2, 2) };
        void Style(string text, string id, double size, string? color)
        {
            var tb = new TextBlock { Text = text, FontSize = size, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center };
            tb.SetResourceReference(TextBlock.ForegroundProperty, color != null ? "AccentTextBrush" : "TextBrush");
            var b = new Button { Content = tb, Width = 96, Height = 62, Margin = new Thickness(2), BorderBrush = (Brush)FindResource("LineBrush") };
            System.Windows.Automation.AutomationProperties.SetName(b, text);
            b.Click += (_, _) => Do(() => _s.SetStyle(id)); gal.Children.Add(b);
        }
        Style("Normale", "Normal", 13, null); Style("Titolo 1", "Heading1", 20, "#0F4761"); Style("Titolo 2", "Heading2", 16, "#0F4761");
        StyleBox = new ComboBox { Width = 140, ToolTip = "Stili", Name = "StyleBox", ItemsSource = BuiltIns.Gallery.Select(x => x.Name).ToList(), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0, 0, 0) };
        StyleBox.DropDownClosed += (_, _) => Style_Changed(null, EventArgs.Empty);
        gal.Children.Add(StyleBox);
        g.Children.Add(gal);

        g = Group(home, "Modifica");
        var ed = new StackPanel();
        Small(ed, "Trova", "find", () => ShowFind(false), "Trova (Ctrl+F)");
        Small(ed, "Sostituisci", "replace", () => ShowFind(true), "Sostituisci (Ctrl+H)");
        Small(ed, "Seleziona tutto", "select", () => Do(() => _s.SelectAll()), "Seleziona tutto (Ctrl+A)");
        g.Children.Add(ed);

        // ====== Inserisci
        var ins = NewTab("Inserisci");
        g = Group(ins, "Pagine");
        var pg = new StackPanel(); Small(pg, "Pagina vuota", "blankpage", () => Do(() => _s.InsertBlankPage())); Small(pg, "Interruzione di pagina", "pagebreak", () => Do(() => _s.InsertPageBreak())); g.Children.Add(pg);
        g = Group(ins, "Tabelle");
        Large(g, "Tabella", "table", () => InsertTable_Click(this, new RoutedEventArgs()));
        g = Group(ins, "Illustrazioni");
        Large(g, "Immagini", "image", () => InsertImage_Click(this, new RoutedEventArgs()));
        var il = new StackPanel(); Stub(il, "Forme", "generic"); Stub(il, "Icone", "generic"); g.Children.Add(il);
        g = Group(ins, "Collegamenti");
        var lk = new StackPanel(); Small(lk, "Collegamento", "generic", () => Link_Click(), "Collegamento: indirizzo web, e-mail o segnalibro (Ctrl+clic per aprirlo)"); Small(lk, "Segnalibro", "generic", () => Bookmark_Click(), "Segnalibro: aggiungi, vai a, elimina"); g.Children.Add(lk);
        g = Group(ins, "Intestazione e piè di pagina");
        var hf = new StackPanel();
        Small(hf, "Intestazione", "header", () => HeaderFooter(true)); Small(hf, "Piè di pagina", "footer", () => HeaderFooter(false));
        Small(hf, "Numero di pagina", "pagenum", () => { var t = _s.GetHeaderFooterText(false); Do(() => _s.SetHeaderFooter(false, t, Align.Center, PageNumFormat.PageOfTotal)); });
        g.Children.Add(hf);

        // ====== Layout
        var lay = NewTab("Layout");
        g = Group(lay, "Imposta pagina", () => PageSetup_Click(this, new RoutedEventArgs()));
        Large(g, "Margini", "margins", () => PageSetup_Click(this, new RoutedEventArgs()));
        Large(g, "Orientamento", "orient", () => Orientation_Click(this, new RoutedEventArgs()));
        Large(g, "Dimensioni", "size", () => PageSetup_Click(this, new RoutedEventArgs()));
        var colBtn = Large(g, "Colonne", "columns", () => { });
        colBtn.Click += (_, _) => Menu(colBtn, ("Una", () => Do(() => _s.SetColumns(1))), ("Due", () => Do(() => _s.SetColumns(2))), ("Tre", () => Do(() => _s.SetColumns(3))));
        var ro = new StackPanel();
        var brk = Small(ro, "Interruzioni", "breaks", () => { });
        brk.Click += (_, _) => Menu(brk, ("Interruzione di pagina", () => Do(() => _s.InsertPageBreak())), ("Sezione: pagina successiva", () => Do(() => _s.InsertSectionBreak(false))), ("Sezione: continua", () => Do(() => _s.InsertSectionBreak(true))));
        Stub(ro, "Numeri di riga", "generic"); g.Children.Add(ro);
        g = Group(lay, "Paragrafo", () => ParaDialog_Click(this, new RoutedEventArgs()));
        TextBox Num(string tip) { var t = new TextBox { Width = 58, Height = 22, ToolTip = tip, VerticalContentAlignment = VerticalAlignment.Center, Margin = new Thickness(2, 1, 2, 1) }; System.Windows.Automation.AutomationProperties.SetName(t, tip); return t; }
        _indL = Num("Rientro a sinistra (cm)"); _indR = Num("Rientro a destra (cm)"); _spB = Num("Spaziatura prima (pt)"); _spA = Num("Spaziatura dopo (pt)");
        void OnEnter(TextBox t, Action<double> apply) => t.KeyDown += (_, e) => { if (e.Key == Key.Enter && double.TryParse(t.Text.Replace("cm", "").Replace("pt", "").Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && v >= 0 && v < 1000) { Do(() => apply(v)); } };
        OnEnter(_indL, v => _s.SetIndent(v * 28.3465, null, null)); OnEnter(_indR, v => _s.SetIndent(null, v * 28.3465, null));
        OnEnter(_spB, v => _s.SetSpacing(v, null)); OnEnter(_spA, v => _s.SetSpacing(null, v));
        TextBlock Lb(string t) => new() { Text = t, Margin = new Thickness(0, 0, 4, 0), VerticalAlignment = VerticalAlignment.Center, Width = 66 };
        g.Children.Add(Col(Row(new TextBlock { Text = "Rientro", FontSize = 12, Foreground = (Brush)FindResource("MutedBrush") }), Row(Lb("A sinistra:"), _indL), Row(Lb("A destra:"), _indR)));
        g.Children.Add(Col(Row(new TextBlock { Text = "Spaziatura", FontSize = 12, Foreground = (Brush)FindResource("MutedBrush") }), Row(Lb("Prima:"), _spB), Row(Lb("Dopo:"), _spA)));

        // ====== Riferimenti
        var rif = NewTab("Riferimenti");
        g = Group(rif, "Note a piè di pagina");
        Large(g, "Inserisci nota a piè di pagina", "footnote", () => InsertFootnote_Click());
        var nt = new StackPanel(); Small(nt, "Inserisci nota di chiusura", "generic", () => InsertEndnote_Click()); Small(nt, "Nota successiva", "generic", () => Do(() => { if (!_s.NextNote()) ThemedDialog.Show(this, "Il documento non contiene note.", "Nota successiva"); }), "Va al richiamo di nota successivo"); g.Children.Add(nt);

        // ====== Visualizza
        var vw = NewTab("Visualizza");
        g = Group(vw, "Visualizzazioni");
        Stub(g, "Modalità di lettura", "read", true);
        var lp = Large(g, "Layout di stampa", "print", () => { }); lp.SetResourceReference(Control.BackgroundProperty, "HoverBrush");
        g = Group(vw, "Mostra");
        var rul = new CheckBox { Content = "Righello", IsChecked = true, Margin = new Thickness(4, 4, 4, 4) };
        rul.Click += (_, _) => { var v = rul.IsChecked == true ? Visibility.Visible : Visibility.Collapsed; HRuler.Visibility = v; VRuler.Visibility = v; };
        var grid = new CheckBox { Content = "Griglia", IsEnabled = false, ToolTip = Soon, Margin = new Thickness(4) };
        var nav = new CheckBox { Content = "Riquadro di spostamento", IsEnabled = false, ToolTip = Soon, Margin = new Thickness(4) };
        g.Children.Add(Col(rul, grid, nav));
        g = Group(vw, "Zoom");
        Large(g, "Zoom", "zoom", () => ZoomDialog());
        Large(g, "100%", "zoom100", () => View.SetZoom(1));
        var zz = new StackPanel();
        Small(zz, "Una pagina", "onepage", () => View.FitPage(Scroller.ViewportWidth, Scroller.ViewportHeight));
        Stub(zz, "Più pagine", "onepage");
        Small(zz, "Larghezza pagina", "pagewidth", () => View.FitWidth(Scroller.ViewportWidth));
        g.Children.Add(zz);

        // ====== Powers
        var pw = NewTab("Powers", power: true);
        g = Group(pw, "Aspetto");
        DarkSwitch = new ToggleButton { IsChecked = !_light, ToolTip = "Modalità scura", Margin = new Thickness(8, 22, 8, 0), VerticalAlignment = VerticalAlignment.Top };
        DarkSwitch.Template = (ControlTemplate)XamlReader.Parse(@"<ControlTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml' TargetType='ToggleButton'>
  <Grid Width='46' Height='24' Background='Transparent'>
    <Border x:Name='tr' CornerRadius='12' BorderThickness='1.5' BorderBrush='#FF2400' Background='Transparent'/>
    <Ellipse x:Name='kn' Width='16' Height='16' Fill='#FF2400' HorizontalAlignment='Left' Margin='4,0,0,0'/>
  </Grid>
  <ControlTemplate.Triggers>
    <Trigger Property='IsChecked' Value='True'><Setter TargetName='tr' Property='Background' Value='#33FF2400'/><Setter TargetName='kn' Property='HorizontalAlignment' Value='Right'/><Setter TargetName='kn' Property='Margin' Value='0,0,4,0'/></Trigger>
  </ControlTemplate.Triggers>
</ControlTemplate>");
        System.Windows.Automation.AutomationProperties.SetName(DarkSwitch, "Modalità scura");
        DarkSwitch.Click += (_, _) => { SetTheme(DarkSwitch.IsChecked != true); _settings.Dark = DarkSwitch.IsChecked == true; _settings.Save(); SelectTab(_tab); };
        var sw = new StackPanel { Orientation = Orientation.Vertical, Width = 130 };
        sw.Children.Add(DarkSwitch);
        sw.Children.Add(new TextBlock { Text = "Modalità scura", HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 8, 0, 0), FontSize = 12 });
        g.Children.Add(sw);

        g = Group(pw, "Strumenti");
        Large(g, "Uniforma", "uniform", () => UniformCommand(), "Uniforma: toglie font, dimensione, colore, evidenziazione e sfondo diretti del testo (selezione o tutto il corpo)");
        Large(g, "Scudo privacy", "shield", () => PrivacyCommand(), "Scudo privacy: trova dati personali nel documento e crea una copia ripulita");

        // ====== Layout di tabella (contestuale)
        var tl = NewTab("Layout di tabella", contextual: true);
        g = Group(tl, "Tabella");
        var t1 = new StackPanel();
        MenuBtn(t1, "Seleziona", "select", false, ("Cella", () => Do(() => _s.SelectTablePart(TableSel.Cell))), ("Riga", () => Do(() => _s.SelectTablePart(TableSel.Row))), ("Colonna", () => Do(() => _s.SelectTablePart(TableSel.Column))), ("Tabella", () => Do(() => _s.SelectTablePart(TableSel.Table))));
        Button? gridBtn = null;
        gridBtn = Small(t1, "Visualizza griglia", "table", () => { View.ShowTableGrid = !View.ShowTableGrid; Mark(gridBtn!, View.ShowTableGrid); View.InvalidateVisual(); }, "Mostra o nasconde i contorni delle celle senza bordi");
        Small(t1, "Proprietà", "table", () => TableProps_Click(), "Proprietà tabella: allineamento, larghezza, altezza righe");
        g.Children.Add(t1);
        g = Group(tl, "Righe e colonne");
        var del = Large(g, "Elimina", "delete", () => { });
        del.Click += (_, _) => Menu(del, ("Elimina riga", () => Do(() => _s.DeleteRow())), ("Elimina tabella", () => Do(() => _s.DeleteTable())));
        Large(g, "Inserisci riga sopra", "rowabove", () => Do(() => _s.AddRowAbove()));
        var rc = new StackPanel();
        Small(rc, "Inserisci riga sotto", "rowbelow", () => Do(() => _s.AddRowBelow()));
        Small(rc, "Inserisci colonna a sinistra", "colleft", () => Do(() => _s.AddColumnLeft()));
        Small(rc, "Inserisci colonna a destra", "colright", () => Do(() => _s.AddColumnRight()));
        g.Children.Add(rc);
        g = Group(tl, "Unione");
        var um = new StackPanel();
        Small(um, "Unisci celle", "merge", () => Do(() => { if (!_s.MergeCells()) ThemedDialog.Show(this, "Seleziona almeno due celle della stessa tabella (trascina con il mouse).", "Unisci celle"); }));
        Small(um, "Dividi celle", "split", () => SplitCells_Click(), "Dividi la cella in più colonne");
        Small(um, "Dividi tabella", "split", () => Do(() => _s.SplitTable()), "Divide la tabella alla riga del cursore");
        g.Children.Add(um);
        g = Group(tl, "Dimensioni cella");
        var dm = new StackPanel();
        MenuBtn(dm, "Adatta", "table", false, ("Adatta al contenuto", () => Do(() => _s.AutoFit(TableFit.Content))), ("Adatta alla finestra", () => Do(() => _s.AutoFit(TableFit.Window))), ("Larghezza colonne fissa", () => Do(() => _s.AutoFit(TableFit.Fixed))));
        Small(dm, "Distribuisci righe", "table", () => Do(() => _s.DistributeRows()), "Rende uguali le righe selezionate (o tutte)");
        Small(dm, "Distribuisci colonne", "table", () => Do(() => _s.DistributeColumns()), "Rende uguali le colonne selezionate (o tutte)");
        g.Children.Add(dm);
        g = Group(tl, "Dati");
        var da = new StackPanel();
        MenuBtn(da, "Ordina", "sort", false, ("Crescente", () => Do(() => _s.SortTable(false))), ("Decrescente", () => Do(() => _s.SortTable(true))));
        Button? hdrBtn = null;
        hdrBtn = Small(da, "Ripeti righe di intestazione", "table", () => { Do(() => _s.ToggleRepeatHeader()); Mark(hdrBtn!, _s.HasRepeatHeader); }, "La prima riga si ripete in cima a ogni pagina");
        Small(da, "Converti in testo", "table", () => Do(() => _s.ConvertTableToText()), "Trasforma la tabella in paragrafi separati da tabulazioni");
        g.Children.Add(da);
    }

    /// <summary>Pulsante con menu a tendina; il nome è cercabile dalla casella Cerca (apre il menu).</summary>
    Button MenuBtn(Panel host, string label, string icon, bool large, params (string Text, Action Act)[] items)
    {
        Button? b = null;
        b = large ? Large(host, label, icon, () => Menu(b!, items)) : Small(host, label, icon, () => Menu(b!, items));
        return b;
    }

    /// <summary>Evidenzia un pulsante a interruttore.</summary>
    static void Mark(Control b, bool on) { if (on) b.SetResourceReference(Control.BackgroundProperty, "HoverBrush"); else b.ClearValue(Control.BackgroundProperty); }

    void Link_Click()
    {
        string cur = _s.HyperlinkAt() ?? "";
        bool sel = _s.HasSelection, fresh = !sel && cur.Length == 0;
        var fields = new List<FormField> { new("Indirizzo (https://… oppure #segnalibro)", cur) };
        if (fresh) fields.Add(new("Testo da visualizzare", ""));
        if (cur.Length > 0) fields.Add(new("Azione", "Modifica", "Modifica", "Rimuovi collegamento"));
        var d = new FormDialog(this, "Collegamento", fields.ToArray());
        if (d.ShowDialog() != true) return;
        var v = d.Values;
        if (cur.Length > 0 && v[^1] == "Rimuovi collegamento") { Do(() => _s.RemoveHyperlink()); return; }
        Do(() => _s.SetHyperlink(v[0], fresh ? v[1] : null));
    }

    void Bookmark_Click()
    {
        var names = _s.Bookmarks();
        string hint = names.Count == 0 ? "" : " (esistenti: " + string.Join(", ", names.Take(6)) + (names.Count > 6 ? "…" : "") + ")";
        var d = new FormDialog(this, "Segnalibro", new FormField("Nome" + hint, names.FirstOrDefault() ?? ""), new FormField("Azione", "Aggiungi", "Aggiungi", "Vai a", "Elimina"));
        if (d.ShowDialog() != true) return;
        var v = d.Values; string name = v[0].Trim();
        switch (v[1])
        {
            case "Vai a": Do(() => { if (!_s.GoToBookmark(name)) throw new InvalidOperationException("Segnalibro non trovato."); }); break;
            case "Elimina": Do(() => { if (!_s.DeleteBookmark(name)) throw new InvalidOperationException("Segnalibro non trovato."); }); break;
            default: Do(() => _s.AddBookmark(name)); break;
        }
    }

    void InsertEndnote_Click()
    {
        var d = new FormDialog(this, "Nota di chiusura", new FormField("Testo della nota", ""));
        if (d.ShowDialog() != true) return;
        Do(() => _s.InsertEndnote(d.Values[0].Trim()));
    }

    /// <summary>Ctrl+clic su un collegamento: segnalibro interno o indirizzo esterno (solo http/https/mailto/ftp, riconvalidato).</summary>
    void OpenLink(string target)
    {
        if (target.Length > 1 && target[0] == '#') { Do(() => _s.GoToBookmark(target[1..])); return; }
        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(EditorSession.NormalizeLink(target)) { UseShellExecute = true }); }
        catch (Exception ex) when (ex is ArgumentException or System.ComponentModel.Win32Exception or InvalidOperationException) { ThemedDialog.Show(this, "Impossibile aprire il collegamento: " + ex.Message, "Collegamento", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    void SplitCells_Click()
    {
        var d = new FormDialog(this, "Dividi celle", new FormField("Numero di colonne (2-12)", "2"));
        if (d.ShowDialog() != true) return;
        var v = d.Numbers();
        if (v == null || v[0] < 2 || v[0] > 12 || v[0] != Math.Floor(v[0])) { ThemedDialog.Show(this, "Colonne da 2 a 12.", "Dividi celle"); return; }
        Do(() => _s.SplitCell((int)v[0]));
    }

    void TableProps_Click()
    {
        EditorSession.TableProps pr;
        try { pr = _s.GetTableProps(); } catch (InvalidOperationException e) { ThemedDialog.Show(this, e.Message, "Proprietà tabella"); return; }
        var d = new FormDialog(this, "Proprietà tabella",
            new FormField("Allineamento", pr.Align, "Sinistra", "Centro", "Destra"),
            new FormField("Larghezza (% del testo, 0 = automatica)", pr.WidthPct.ToString()),
            new FormField("Altezza minima righe (pt, 0 = automatica)", F(pr.RowHeightPt)));
        if (d.ShowDialog() != true) return;
        var v = d.Values;
        if (!double.TryParse(v[1].Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var w) || w < 0 || w > 100 || !double.TryParse(v[2].Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var h) || h < 0 || h > 1584)
        { ThemedDialog.Show(this, "Larghezza da 0 a 100, altezza da 0 a 1584.", "Proprietà tabella"); return; }
        Do(() => _s.SetTableProps(v[0], (int)Math.Round(w), h));
    }

    void Menu(FrameworkElement anchor, params (string Text, Action Act)[] items)
    {
        var m = new ContextMenu { PlacementTarget = anchor, Placement = PlacementMode.Bottom };
        foreach (var (t, a) in items) { var it = new MenuItem { Header = t }; var aa = a; it.Click += (_, _) => aa(); m.Items.Add(it); }
        m.IsOpen = true;
    }

    // ---------------------------------------------------------------- Powers
    public void RunUniformPreview() => UniformCommand();
    public void RunPrivacy() => PrivacyCommand();

    void UniformCommand()
    {
        var pv = _s.PreviewUniform();
        string scope = _s.HasSelection ? "nella selezione" : "nel corpo del documento";
        if (pv.Runs == 0) { ThemedDialog.Show(this, $"Niente da uniformare {scope}." + (pv.Left > 0 ? $" {pv.Left} paragrafi simili a titoli, senza stile, restano come sono." : ""), "Uniforma", MessageBoxButton.OK, MessageBoxImage.Information); return; }
        string msg = $"{pv.Runs} modifiche in {pv.Paragraphs} paragrafi {scope}: font, dimensione, colore, evidenziazione e sfondo del testo tornano a quelli dello stile. Restano grassetto, corsivo, sottolineato, apice/pedice, collegamenti, immagini, tabelle ed elenchi."
            + (pv.Left > 0 ? Environment.NewLine + $"{pv.Left} paragrafi senza stile e simili a titoli restano come sono." : "");
        if (ThemedDialog.Confirm(this, "Uniforma", msg, "Applica")) Do(() => _s.Uniform());
    }

    void PrivacyCommand()
    {
        byte[] bytes;
        try { bytes = _doc.ToBytes(); }
        catch (Exception e) when (e is System.IO.IOException or InvalidOperationException) { ThemedDialog.Show(this, "Impossibile leggere il documento: " + e.Message, "Scudo privacy", MessageBoxButton.OK, MessageBoxImage.Error); return; }
        PrivacyDialog.Show(this, bytes, _doc.Path, DocDisplayName);
    }

    // ---------------------------------------------------------------- finestre
    void ColorMenu() => Color_Click(this, new RoutedEventArgs());
    void HighlightMenu() => Highlight_Click(this, new RoutedEventArgs());

    void ZoomDialog()
    {
        var d = new FormDialog(this, "Zoom", new FormField("Percentuale (25-400)", ((int)Math.Round(View.Zoom * 100)).ToString()));
        if (d.ShowDialog() != true) return;
        var v = d.Numbers();
        if (v == null || v[0] < 25 || v[0] > 400) { ThemedDialog.Show(this, "Valore da 25 a 400.", "Zoom"); return; }
        View.SetZoom(v[0] / 100);
    }

    void CharDialog()
    {
        var f = _s.FormatAtCaret();
        string style = f.Bold && f.Italic ? "Grassetto corsivo" : f.Bold ? "Grassetto" : f.Italic ? "Corsivo" : "Normale";
        var d = new FormDialog(this, "Carattere",
            new FormField("Tipo di carattere", f.Font), new FormField("Stile", style, "Normale", "Corsivo", "Grassetto", "Grassetto corsivo"),
            new FormField("Dimensione (pt)", f.Size.ToString("0.#", CultureInfo.CurrentCulture)),
            new FormField("Colore (RRGGBB o vuoto = automatico)", f.Color ?? ""),
            new FormField("Sottolineato", f.Underline ? "Sì" : "No", "No", "Sì"), new FormField("Barrato", f.Strike ? "Sì" : "No", "No", "Sì"),
            new FormField("Apice/pedice", f.VertAlign > 0 ? "Apice" : f.VertAlign < 0 ? "Pedice" : "Nessuno", "Nessuno", "Apice", "Pedice"),
            new FormField("Maiuscole", f.AllCaps ? "Tutte maiuscole" : f.SmallCaps ? "Maiuscoletto" : "Normale", "Normale", "Tutte maiuscole", "Maiuscoletto"));
        if (d.ShowDialog() != true) return;
        var v = d.Values;
        if (!double.TryParse(v[2].Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var size) || size < 1 || size > 1638) { ThemedDialog.Show(this, "Dimensione non valida (1-1638).", "Carattere"); return; }
        var col = v[3].Trim().TrimStart('#');
        if (col.Length != 0 && (col.Length != 6 || !col.All(Uri.IsHexDigit))) { ThemedDialog.Show(this, "Colore non valido: usa 6 cifre esadecimali.", "Carattere"); return; }
        Do(() => _s.Batch(() =>
        {
            _s.SetFont(v[0].Trim()); _s.SetSize(size);
            bool b = v[1].StartsWith("Grassetto"), i = v[1].EndsWith("corsivo") || v[1] == "Corsivo";
            _s.ApplyRun(r => { r.Bold = new DocumentFormat.OpenXml.Wordprocessing.Bold { Val = b }; r.Italic = new DocumentFormat.OpenXml.Wordprocessing.Italic { Val = i };
                r.Underline = new DocumentFormat.OpenXml.Wordprocessing.Underline { Val = v[4] == "Sì" ? DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single : DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.None };
                r.Strike = new DocumentFormat.OpenXml.Wordprocessing.Strike { Val = v[5] == "Sì" }; });
            _s.SetVertAlign(v[6] == "Apice" ? 1 : v[6] == "Pedice" ? -1 : 0);
            _s.SetCaps(v[7] != "Normale", v[7] == "Maiuscoletto");
            _s.SetColor(col.Length == 0 ? null : col.ToUpperInvariant());
        }));
    }

    // ---------------------------------------------------------------- stato
    void BuildStatusBar()
    {
        Button V(string icon, bool on, bool enabled) { var b = new Button { Content = Icons.Make(icon, 16), Padding = new Thickness(5, 2, 5, 2), IsEnabled = enabled, ToolTip = enabled ? "Layout di stampa" : Soon }; if (on) { b.SetResourceReference(Control.BackgroundProperty, "HoverBrush"); _printViewBtn = b; } ViewBar.Children.Add(b); return b; }
        V("read", false, false); V("print", true, true);
        var minus = new Button { Content = Icons.Make("close", 10), Padding = new Thickness(4), ToolTip = "Riduci zoom" };
        minus.Content = new TextBlock { Text = "−" }; minus.Click += (_, _) => View.SetZoom(View.Zoom / 1.1);
        var plus = new Button { Content = new TextBlock { Text = "+" }, Padding = new Thickness(4), ToolTip = "Aumenta zoom" }; plus.Click += (_, _) => View.SetZoom(View.Zoom * 1.1);
        _zoomSlider = new Slider { Minimum = 25, Maximum = 400, Value = 100, Width = 120, VerticalAlignment = VerticalAlignment.Center, ToolTip = "Zoom" };
        _zoomSlider.ValueChanged += (_, e) => { if (!_updating && Math.Abs(View.Zoom * 100 - e.NewValue) > 0.5) View.SetZoom(e.NewValue / 100); };
        _zoomPct = new TextBlock { Width = 44, TextAlignment = TextAlignment.Right, VerticalAlignment = VerticalAlignment.Center, Text = "100%" };
        ZoomBar.Children.Add(minus); ZoomBar.Children.Add(_zoomSlider); ZoomBar.Children.Add(plus); ZoomBar.Children.Add(_zoomPct);
    }

    void UpdateZoomUi()
    {
        _zoomPct.Text = $"{View.Zoom * 100:0}%";
        _updating = true; _zoomSlider.Value = Math.Clamp(View.Zoom * 100, 25, 400); _updating = false;
    }

    void UpdateRulers()
    {
        if (_s == null) return;
        var g = View.CaretPageGeometry();
        if (g == null) return;
        var (r, pg, scale) = g.Value;
        var origin = View.TranslatePoint(new Point(0, 0), Scroller);   // vista → viewport
        double leftDip = r.X + origin.X, topDip = r.Y + origin.Y;
        HRuler.Set(leftDip, scale, pg.Width, pg.MarL, pg.MarR);
        VRuler.Set(topDip, scale, pg.Height, pg.MarT, pg.MarB);
        // aggiorna anche le caselle rientro/spaziatura se non in uso
        try
        {
            var pf = _s.ParaFormatAtCaret();
            if (_indL != null && !_indL.IsKeyboardFocusWithin) _indL.Text = (pf.IndLeft / 28.3465).ToString("0.##", CultureInfo.CurrentCulture) + " cm";
            if (_indR != null && !_indR.IsKeyboardFocusWithin) _indR.Text = (pf.IndRight / 28.3465).ToString("0.##", CultureInfo.CurrentCulture) + " cm";
            if (_spB != null && !_spB.IsKeyboardFocusWithin) _spB.Text = pf.Before.ToString("0.##", CultureInfo.CurrentCulture) + " pt";
            if (_spA != null && !_spA.IsKeyboardFocusWithin) _spA.Text = pf.After.ToString("0.##", CultureInfo.CurrentCulture) + " pt";
        }
        catch (InvalidOperationException) { }
    }

    // ---------------------------------------------------------------- scheda File
    void BuildFilePage()
    {
        FileNav.Children.Clear();
        Button Item(string text, string icon, Action a)
        {
            var b = new Button { Content = new StackPanel { Orientation = Orientation.Horizontal, Children = { Icons.Make(icon, 20), new TextBlock { Text = text, Margin = new Thickness(12, 0, 0, 0), FontSize = 14, VerticalAlignment = VerticalAlignment.Center } } }, HorizontalContentAlignment = HorizontalAlignment.Left, Padding = new Thickness(14, 9, 14, 9) };
            System.Windows.Automation.AutomationProperties.SetName(b, text);
            b.Click += (_, _) => a(); FileNav.Children.Add(b); return b;
        }
        var back = Item("Indietro", "caret", HideFile);
        Item("Nuovo", "new", () => { HideFile(); New_Click(this, new RoutedEventArgs()); });
        Item("Apri", "open", () => { HideFile(); Open_Click(this, new RoutedEventArgs()); });
        Item("Salva", "save", () => { HideFile(); SaveCore(false); });
        Item("Salva con nome", "save", () => { HideFile(); SaveCore(true); });
        Item("Esporta in PDF", "export", () => { HideFile(); Pdf_Click(this, new RoutedEventArgs()); });
        Item("Stampa", "print", () => { HideFile(); Print_Click(this, new RoutedEventArgs()); });
        Item("Chiudi", "close", () => { HideFile(); Close(); });
    }

    void ShowFile(string title)
    {
        FileTitle.Text = title;
        FileInfo.Text = _doc.Path == null ? "Documento non ancora salvato." : $"{_doc.Path}\n{View.PageCount} pagine, {_wc} parole.";
        FilePage.Visibility = Visibility.Visible;
    }

    void HideFile() { FilePage.Visibility = Visibility.Collapsed; View.Focus(); }
}
