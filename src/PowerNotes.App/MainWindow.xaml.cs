using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using HighlightColorValues = DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues;
using Microsoft.Win32;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Model;
using PowerNotes.Core.Pdf;
using PowerNotes.Core.Settings;

namespace PowerNotes.App;

public partial class MainWindow : Window
{
    WordDoc _doc = null!;
    EditorSession _s = null!;
    bool _updating;
    bool _light;
    AppSettings _settings = new();
    FindDialog? _find;
    static readonly string[] Palette = ["000000", "7F7F7F", "C00000", "FF0000", "FFC000", "FFFF00", "00B050", "00B0F0", "0070C0", "002060", "7030A0", "FFFFFF"];
    static readonly (string Name, HighlightColorValues V)[] Highlights =
    [("Giallo", HighlightColorValues.Yellow), ("Verde", HighlightColorValues.Green), ("Azzurro", HighlightColorValues.Cyan), ("Rosa", HighlightColorValues.Magenta), ("Rosso", HighlightColorValues.Red), ("Grigio", HighlightColorValues.LightGray)];

    public MainWindow()
    {
        InitializeComponent();
        View.Scroller = Scroller;
        View.StateChanged += UpdateUi;
        Scroller.ScrollChanged += (_, _) => View.InvalidateVisual();
        SizeChanged += (_, _) => View.InvalidateVisual();
        BuildUi();
        PreviewKeyDown += OnShortcut;
        Closing += (_, e) => { if (!ConfirmDiscard()) e.Cancel = true; };
        _settings = AppSettings.Load();
        SetTheme(!_settings.Dark);
        Load(WordDoc.CreateNew());
        Loaded += (_, _) => View.Focus();
    }

    public EditorSession Session => _s;
    public WordDoc Doc => _doc;

    // ---------- documento ----------
    void Load(WordDoc d)
    {
        _doc?.Dispose();
        _doc = d; _s = new EditorSession(d); _wcVersion = -1;
        View.Session = _s;
        UpdateTitle(); UpdateUi();
        Scroller.ScrollToTop();
    }

    public void OpenPath(string path)
    {
        try { Load(WordDoc.Open(path)); }
        catch (DocumentLoadException e) { MessageBox.Show(this, e.Message, "Impossibile aprire il documento", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    bool ConfirmDiscard()
    {
        if (_doc == null || !_doc.Dirty) return true;
        var r = MessageBox.Show(this, "Il documento ha modifiche non salvate. Salvarle?", "PowerNotes", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (r == MessageBoxResult.Cancel) return false;
        return r == MessageBoxResult.No || SaveCore(false);
    }

    void UpdateTitle()
    {
        var name = _doc.Path == null ? "Documento senza nome" : Path.GetFileName(_doc.Path);
        Title = $"{name}{(_doc.Dirty ? " •" : "")} — PowerNotes";
        TitleLabel.Text = $"{Path.GetFileNameWithoutExtension(name)}{(_doc.Dirty ? " •" : "")}  -  PowerNotes";
    }

    bool SaveCore(bool saveAs)
    {
        var path = _doc.Path;
        if (saveAs || path == null)
        {
            var dlg = new SaveFileDialog { Filter = "Documento Word (*.docx)|*.docx", DefaultExt = ".docx", FileName = path == null ? "Documento.docx" : Path.GetFileName(path) };
            if (dlg.ShowDialog(this) != true) return false;
            path = dlg.FileName;
        }
        try { _doc.Save(path); UpdateTitle(); return true; }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or InvalidOperationException or DocumentLoadException or DocumentFormat.OpenXml.Packaging.OpenXmlPackageException)
        {
            MessageBox.Show(this, "Salvataggio non riuscito, il file originale non è stato modificato.\n\n" + e.Message, "Errore", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    // ---------- menu File ----------
    void New_Click(object s, RoutedEventArgs e) { if (ConfirmDiscard()) Load(WordDoc.CreateNew()); }
    void Open_Click(object s, RoutedEventArgs e)
    {
        if (!ConfirmDiscard()) return;
        var dlg = new OpenFileDialog { Filter = "Documenti Word (*.docx)|*.docx" };
        if (dlg.ShowDialog(this) == true) OpenPath(dlg.FileName);
    }
    void Save_Click(object s, RoutedEventArgs e) => SaveCore(false);
    void SaveAs_Click(object s, RoutedEventArgs e) => SaveCore(true);
    void Exit_Click(object s, RoutedEventArgs e) => Close();

    void Pdf_Click(object s, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", DefaultExt = ".pdf", FileName = Path.GetFileNameWithoutExtension(_doc.Path ?? "Documento") + ".pdf" };
        if (dlg.ShowDialog(this) != true) return;
        try { File.WriteAllBytes(dlg.FileName, PdfExporter.Export(_s.Layout, Path.GetFileNameWithoutExtension(dlg.FileName))); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException) { MessageBox.Show(this, "Esportazione PDF non riuscita: " + ex.Message, "Errore", MessageBoxButton.OK, MessageBoxImage.Error); }
    }

    void Print_Click(object s, RoutedEventArgs e)
    {
        var pd = new PrintDialog { UserPageRangeEnabled = true, MinPage = 1, MaxPage = (uint)Math.Max(1, View.PageCount) };
        if (pd.ShowDialog() != true) return;
        try { pd.PrintDocument(new LayoutPaginator(_s.Layout), Path.GetFileName(_doc.Path ?? "Documento")); }
        catch (Exception ex) when (ex is InvalidOperationException or System.Printing.PrintQueueException) { MessageBox.Show(this, "Stampa non riuscita: " + ex.Message, "Errore", MessageBoxButton.OK, MessageBoxImage.Error); }
    }

    // ---------- scorciatoie ----------
    void OnShortcut(object sender, KeyEventArgs e)
    {
        if (!Keyboard.Modifiers.HasFlag(ModifierKeys.Control) || Keyboard.Modifiers.HasFlag(ModifierKeys.Alt)) return;
        if (Keyboard.FocusedElement is TextBox && !View.IsKeyboardFocused) return;
        bool shift = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift);
        Action? a = e.Key switch
        {
            Key.B => () => _s.ToggleBold(), Key.I => () => _s.ToggleItalic(), Key.U => () => _s.ToggleUnderline(),
            Key.Z => () => _s.Undo(), Key.Y => () => _s.Redo(), Key.A => () => _s.SelectAll(),
            Key.C => () => Copy(), Key.X => () => Cut(), Key.V => () => Paste(),
            Key.L => () => _s.SetAlignment(Align.Left), Key.E => () => _s.SetAlignment(Align.Center), Key.R => () => _s.SetAlignment(Align.Right), Key.J => () => _s.SetAlignment(Align.Justify),
            Key.S => () => SaveCore(shift), Key.O => () => Open_Click(this, e), Key.N => () => New_Click(this, e), Key.P => () => Print_Click(this, e),
            Key.F => () => Find_Click(this, e), Key.H => () => Replace_Click(this, e),
            Key.Enter => () => { _s.InsertPageBreak(); },
            Key.OemPlus or Key.Add => () => View.SetZoom(View.Zoom * 1.1), Key.OemMinus or Key.Subtract => () => View.SetZoom(View.Zoom / 1.1), Key.D0 or Key.NumPad0 => () => View.SetZoom(1),
            _ => null
        };
        if (a == null) return;
        e.Handled = true; Do(a);
    }

    /// <summary>Esegue un comando di modifica, gestendo errori utente senza crash, e aggiorna la vista.</summary>
    void Do(Action a)
    {
        try { a(); }
        catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or IOException or UnauthorizedAccessException)
        { MessageBox.Show(this, ex.Message, "PowerNotes", MessageBoxButton.OK, MessageBoxImage.Warning); }
        View.Refresh(true);
        View.Focus();
    }

    // ---------- appunti ----------
    const string FragFormat = "PowerNotes.Fragment";

    /// <summary>Appunti: testo semplice + frammento formattato (formattazione conservata fra documenti di PowerNotes).</summary>
    void Copy()
    {
        if (!_s.HasSelection) return;
        var data = new DataObject();
        data.SetText(_s.SelectedText().Replace(((char)10).ToString(), ((char)13).ToString() + (char)10));
        if (_s.CopyFragment() is { } frag) data.SetData(FragFormat, _doc.Id + "|" + frag);
        Clipboard.SetDataObject(data, true);
    }
    void Cut() { if (!_s.HasSelection) return; Copy(); _s.DeleteSelection(); }
    void Paste()
    {
        if (Clipboard.GetData(FragFormat) is string f && f.IndexOf('|') is var i and > 0)
        {
            try { _s.PasteFragment(f[(i + 1)..], f[..i] == _doc.Id.ToString()); return; }
            catch (InvalidOperationException) { /* frammento non valido: ripiego sul testo */ }
        }
        if (Clipboard.ContainsText()) _s.Type(Clipboard.GetText());
    }
    void Copy_Click(object s, RoutedEventArgs e) => Do(Copy);
    void Cut_Click(object s, RoutedEventArgs e) => Do(Cut);
    void Paste_Click(object s, RoutedEventArgs e) => Do(Paste);
    void Undo_Click(object s, RoutedEventArgs e) => Do(() => _s.Undo());
    void Redo_Click(object s, RoutedEventArgs e) => Do(() => _s.Redo());

    // ---------- carattere ----------
    void Font_Changed(object? s, EventArgs e) { if (_updating || string.IsNullOrWhiteSpace(FontBox.Text)) return; var n = FontBox.Text.Trim(); Do(() => _s.SetFont(n)); }
    void Size_Changed(object? s, EventArgs e)
    {
        if (_updating) return;
        if (double.TryParse(SizeBox.Text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && v >= 1 && v <= 1638) Do(() => _s.SetSize(v));
    }
    void Bold_Click(object s, RoutedEventArgs e) => Do(() => _s.ToggleBold());
    void Italic_Click(object s, RoutedEventArgs e) => Do(() => _s.ToggleItalic());
    void Under_Click(object s, RoutedEventArgs e) => Do(() => _s.ToggleUnderline());
    void Strike_Click(object s, RoutedEventArgs e) => Do(() => _s.ToggleStrike());
    void Super_Click(object s, RoutedEventArgs e) => Do(() => _s.SetVertAlign(_s.AllRuns(f => f.VertAlign > 0) ? 0 : 1));
    void Sub_Click(object s, RoutedEventArgs e) => Do(() => _s.SetVertAlign(_s.AllRuns(f => f.VertAlign < 0) ? 0 : -1));

    void Color_Click(object s, RoutedEventArgs e)
    {
        var m = new ContextMenu();
        foreach (var hex in Palette)
        {
            var h = hex;
            var item = new MenuItem { Header = new Border { Width = 40, Height = 14, Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#" + h)), BorderBrush = Brushes.Gray, BorderThickness = new Thickness(1) } };
            item.Click += (_, _) => Do(() => _s.SetColor(h));
            m.Items.Add(item);
        }
        var auto = new MenuItem { Header = "Automatico" }; auto.Click += (_, _) => Do(() => _s.SetColor(null)); m.Items.Add(auto);
        m.PlacementTarget = (UIElement)s; m.IsOpen = true;
    }

    void Highlight_Click(object s, RoutedEventArgs e)
    {
        var m = new ContextMenu();
        foreach (var (name, v) in Highlights) { var vv = v; var it = new MenuItem { Header = name }; it.Click += (_, _) => Do(() => _s.SetHighlight(vv)); m.Items.Add(it); }
        var none = new MenuItem { Header = "Nessuna" }; none.Click += (_, _) => Do(() => _s.SetHighlight(null)); m.Items.Add(none);
        m.PlacementTarget = (UIElement)s; m.IsOpen = true;
    }

    // ---------- paragrafo ----------
    void AlLeft_Click(object s, RoutedEventArgs e) => Do(() => _s.SetAlignment(Align.Left));
    void AlCenter_Click(object s, RoutedEventArgs e) => Do(() => _s.SetAlignment(Align.Center));
    void AlRight_Click(object s, RoutedEventArgs e) => Do(() => _s.SetAlignment(Align.Right));
    void AlJust_Click(object s, RoutedEventArgs e) => Do(() => _s.SetAlignment(Align.Justify));
    void Line_Changed(object? s, EventArgs e)
    {
        if (_updating || LineBox.SelectedItem is not string t) return;
        if (double.TryParse(t.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) Do(() => _s.SetLineSpacing(v));
    }
    void Bullets_Click(object s, RoutedEventArgs e) => Do(() => _s.ToggleList("bullet"));
    void Numbers_Click(object s, RoutedEventArgs e) => Do(() => _s.ToggleList("decimal"));
    void OutIndent_Click(object s, RoutedEventArgs e) => Do(() => _s.ChangeIndent(-36));
    void InIndent_Click(object s, RoutedEventArgs e) => Do(() => _s.ChangeIndent(36));
    void Style_Changed(object? s, EventArgs e)
    {
        if (_updating || StyleBox.SelectedIndex < 0) return;
        var id = BuiltIns.Gallery[StyleBox.SelectedIndex].Id;
        Do(() => _s.SetStyle(id));
    }

    void ParaDialog_Click(object s, RoutedEventArgs e)
    {
        var f = _s.ParaFormatAtCaret();
        var d = new FormDialog(this, "Paragrafo",
            new FormField("Rientro sinistro (pt)", F(f.IndLeft)), new FormField("Rientro destro (pt)", F(f.IndRight)),
            new FormField("Prima riga (pt, negativo = sporgente)", F(f.FirstLine)),
            new FormField("Spazio prima (pt)", F(f.Before)), new FormField("Spazio dopo (pt)", F(f.After)),
            new FormField("Interlinea (multiplo)", F(f.LineRule == "auto" ? f.Line / 240 : 1)));
        if (d.ShowDialog() != true) return;
        var v = d.Numbers();
        if (v == null) { MessageBox.Show(this, "Valori numerici non validi.", "Paragrafo"); return; }
        Do(() => { _s.SetIndent(v[0], v[1], v[2]); _s.SetSpacing(v[3], v[4]); if (v[5] >= 0.5 && v[5] <= 9) _s.SetLineSpacing(v[5]); });
    }

    static string F(double d) => d.ToString("0.##", CultureInfo.CurrentCulture);

    // ---------- trova/sostituisci ----------
    void Find_Click(object s, RoutedEventArgs e) => ShowFind(false);
    void Replace_Click(object s, RoutedEventArgs e) => ShowFind(true);
    void ShowFind(bool replace)
    {
        if (_find == null || !_find.IsLoaded) { _find = new FindDialog(this, _s, () => { View.Refresh(true); }); _find.Show(); }
        _find.SetMode(replace);
        _find.Activate();
    }

    // ---------- inserisci ----------
    void InsertTable_Click(object s, RoutedEventArgs e)
    {
        var d = new FormDialog(this, "Inserisci tabella", new FormField("Righe", "3"), new FormField("Colonne", "3"));
        if (d.ShowDialog() != true) return;
        var v = d.Numbers();
        if (v == null || v[0] < 1 || v[1] < 1 || v[0] > 200 || v[1] > 30) { MessageBox.Show(this, "Righe da 1 a 200, colonne da 1 a 30.", "Tabella"); return; }
        Do(() => _s.InsertTable((int)v[0], (int)v[1]));
    }

    void InsertImage_Click(object s, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "Immagini|*.png;*.jpg;*.jpeg;*.gif;*.bmp" };
        if (dlg.ShowDialog(this) == true) Do(() => _s.InsertImage(dlg.FileName));
    }

    void PageBreak_Click(object s, RoutedEventArgs e) => Do(() => _s.InsertPageBreak());
    void SoftBreak_Click(object s, RoutedEventArgs e) => Do(() => _s.SoftBreak());
    void Header_Click(object s, RoutedEventArgs e) => HeaderFooter(true);
    void Footer_Click(object s, RoutedEventArgs e) => HeaderFooter(false);
    void PageNumbers_Click(object s, RoutedEventArgs e)
    {
        var text = _s.GetHeaderFooterText(false);
        Do(() => _s.SetHeaderFooter(false, text, Align.Center, PageNumFormat.PageOfTotal));
    }

    void HeaderFooter(bool header)
    {
        var d = new FormDialog(this, header ? "Intestazione" : "Piè di pagina",
            new FormField("Testo", _s.GetHeaderFooterText(header)),
            new FormField("Allineamento", "Centro", "Sinistra", "Centro", "Destra"),
            new FormField("Numero di pagina", "Nessuno", "Nessuno", "Numero", "Pagina X di Y"));
        if (d.ShowDialog() != true) return;
        var r = d.Values;
        var al = r[1] == "Sinistra" ? Align.Left : r[1] == "Destra" ? Align.Right : Align.Center;
        var pn = r[2] == "Numero" ? PageNumFormat.Number : r[2] == "Pagina X di Y" ? PageNumFormat.PageOfTotal : PageNumFormat.None;
        Do(() => _s.SetHeaderFooter(header, r[0], al, pn));
    }

    // ---------- layout ----------
    void PageSetup_Click(object s, RoutedEventArgs e)
    {
        var p = _s.GetPageSetup();
        const double Cm = 28.3465;
        string paper = Math.Abs(Math.Min(p.Width, p.Height) - 595.3) < 3 ? "A4" : Math.Abs(Math.Min(p.Width, p.Height) - 612) < 3 ? "Letter" : Math.Abs(Math.Min(p.Width, p.Height) - 419.5) < 3 ? "A5" : Math.Abs(Math.Min(p.Width, p.Height) - 612) < 3 ? "Letter" : "Personalizzato";
        var d = new FormDialog(this, "Imposta pagina",
            new FormField("Formato", paper == "Personalizzato" ? "A4" : paper, "A4", "A5", "Letter", "Legal"),
            new FormField("Orientamento", p.Landscape ? "Orizzontale" : "Verticale", "Verticale", "Orizzontale"),
            new FormField("Margine superiore (cm)", F(p.Top / Cm)), new FormField("Margine inferiore (cm)", F(p.Bottom / Cm)),
            new FormField("Margine sinistro (cm)", F(p.Left / Cm)), new FormField("Margine destro (cm)", F(p.Right / Cm)));
        if (d.ShowDialog() != true) return;
        var v = d.Values;
        double w, h;
        switch (v[0]) { case "A5": (w, h) = (419.5, 595.3); break; case "Letter": (w, h) = (612, 792); break; case "Legal": (w, h) = (612, 1008); break; default: (w, h) = (595.3, 841.9); break; }
        bool land = v[1] == "Orizzontale";
        var m = new double[4];
        for (int i = 0; i < 4; i++)
            if (!double.TryParse(v[2 + i].Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out m[i]) || m[i] < 0) { MessageBox.Show(this, "Margini non validi.", "Imposta pagina"); return; }
        Do(() => _s.SetPageSetup(new EditorSession.PageInfo(w, h, land, m[0] * Cm, m[1] * Cm, m[2] * Cm, m[3] * Cm)));
    }

    void Orientation_Click(object s, RoutedEventArgs e)
    {
        var p = _s.GetPageSetup();
        Do(() => _s.SetPageSetup(p with { Landscape = !p.Landscape }));
    }

    // ---------- tabella ----------
    void MergeCells_Click(object s, RoutedEventArgs e) => Do(() => { if (!_s.MergeCells()) MessageBox.Show(this, "Seleziona almeno due celle della stessa tabella (trascina con il mouse).", "Unisci celle"); });
    void AddRow_Click(object s, RoutedEventArgs e) => Do(() => _s.AddRowBelow());
    void AddCol_Click(object s, RoutedEventArgs e) => Do(() => _s.AddColumnRight());
    void DelRow_Click(object s, RoutedEventArgs e) => Do(() => _s.DeleteRow());
    void DelTable_Click(object s, RoutedEventArgs e) => Do(() => _s.DeleteTable());

    // ---------- vista ----------
    void ZoomIn_Click(object s, RoutedEventArgs e) => View.SetZoom(View.Zoom * 1.15);
    void ZoomOut_Click(object s, RoutedEventArgs e) => View.SetZoom(View.Zoom / 1.15);
    void Zoom100_Click(object s, RoutedEventArgs e) => View.SetZoom(1);
    void Theme_Click(object s, RoutedEventArgs e) => SetTheme(!_light);

    public void SetTheme(bool light)
    {
        _light = light;
        Application.Current.Resources.MergedDictionaries[0] = new ResourceDictionary { Source = new Uri(light ? "Themes/Light.xaml" : "Themes/Dark.xaml", UriKind.Relative) };
        View.PaperShadow = new SolidColorBrush(Color.FromArgb(light ? (byte)40 : (byte)90, 0, 0, 0));
        View.InvalidateVisual(); HRuler?.InvalidateVisual(); VRuler?.InvalidateVisual();
        if (_printViewBtn != null) _printViewBtn.SetResourceReference(Control.BackgroundProperty, "HoverBrush");
        if (DarkSwitch != null) DarkSwitch.IsChecked = !light;
    }

    public bool IsLight => _light;

    /// <summary>Solo per le schermate di documentazione: seleziona i primi n caratteri del primo paragrafo non vuoto.</summary>
    public void ApplyDemoSelection(int n)
    {
        var p = _s.Paragraphs().FirstOrDefault(x => ParaText.Length(x) > 0);
        if (p == null) return;
        _s.SetCaret(p, 0); _s.SetCaret(p, Math.Min(n, ParaText.Length(p)), true);
        View.Refresh();
    }

    // ---------- stato ----------
    int _wcVersion = -1, _wc;

    void UpdateUi()
    {
        if (_s == null) return;
        _updating = true;
        try
        {
            var f = _s.FormatAtCaret();
            var pf = _s.ParaFormatAtCaret();
            BoldBtn.IsChecked = _s.AllRuns(x => x.Bold); ItalicBtn.IsChecked = _s.AllRuns(x => x.Italic);
            UnderBtn.IsChecked = _s.AllRuns(x => x.Underline); StrikeBtn.IsChecked = _s.AllRuns(x => x.Strike);
            AlLeft.IsChecked = pf.Alignment == Align.Left; AlCenter.IsChecked = pf.Alignment == Align.Center;
            AlRight.IsChecked = pf.Alignment == Align.Right; AlJust.IsChecked = pf.Alignment == Align.Justify;
            if (!FontBox.IsKeyboardFocusWithin) FontBox.Text = f.Font;
            if (!SizeBox.IsKeyboardFocusWithin) SizeBox.Text = f.Size.ToString("0.#", CultureInfo.CurrentCulture);
            var sid = _s.Styles.Find(pf.StyleId);
            var gi = Array.FindIndex(BuiltIns.Gallery, g => string.Equals(g.WordName, sid?.StyleName?.Val?.Value, StringComparison.OrdinalIgnoreCase) || (g.Id == "Normal" && sid?.Default?.Value == true));
            StyleBox.SelectedIndex = gi >= 0 ? gi : (pf.StyleId == null ? 0 : -1);
            SetTableTab(_s.InTable);
            UpdateRulers();
            PageText.Text = $"Pagina {View.CaretPage} di {View.PageCount}";
            if (_wcVersion != _doc.Version) { _wcVersion = _doc.Version; _wc = _s.Paragraphs().Sum(p => ParaText.Text(p).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length); }
            WordsText.Text = $"{_wc} parole";
            UpdateZoomUi();
            var lm = pf.LineRule == "auto" ? pf.Line / 240 : 0; LineBox.SelectedItem = ((string[])LineBox.ItemsSource).OrderBy(x => Math.Abs(double.Parse(x.Replace(',', '.'), CultureInfo.InvariantCulture) - lm)).First() is var near && Math.Abs(double.Parse(near.Replace(',', '.'), CultureInfo.InvariantCulture) - lm) < 0.02 ? near : null;
            UpdateTitle();
        }
        catch (Exception ex) when (ex is InvalidOperationException or KeyNotFoundException or ArgumentException) { PageText.Text = "Errore di visualizzazione: " + ex.Message; }
        finally { _updating = false; }
    }
}
