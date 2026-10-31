using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using PowerNotes.Core.Edit;
using PowerNotes.Core.Layout;

namespace PowerNotes.App;

public sealed record FormField(string Label, string Value, params string[] Options);

/// <summary>Finestra modale a campi (testo o scelta) costruita da codice, nel tema corrente.</summary>
public sealed class FormDialog : Window
{
    readonly List<Control> _inputs = new();

    public FormDialog(Window owner, string title, params FormField[] fields)
    {
        Owner = owner; Title = title; SizeToContent = SizeToContent.WidthAndHeight; ResizeMode = ResizeMode.NoResize;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var grid = new Grid { Margin = new Thickness(16) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });
        int r = 0;
        foreach (var f in fields)
        {
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var l = new TextBlock { Text = f.Label, Margin = new Thickness(0, 6, 14, 6), VerticalAlignment = VerticalAlignment.Center };
            Grid.SetRow(l, r); grid.Children.Add(l);
            Control c;
            if (f.Options.Length > 0) { var cb = new ComboBox { ItemsSource = f.Options, SelectedItem = f.Value, Margin = new Thickness(0, 4, 0, 4) }; if (cb.SelectedItem == null) cb.SelectedIndex = 0; c = cb; }
            else c = new TextBox { Text = f.Value, Margin = new Thickness(0, 4, 0, 4), Height = 26, VerticalContentAlignment = VerticalAlignment.Center };
            Grid.SetRow(c, r); Grid.SetColumn(c, 1); grid.Children.Add(c);
            _inputs.Add(c); r++;
        }
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        var ok = new Button { Content = "OK", IsDefault = true, Width = 80, BorderBrush = (Brush)FindResource("AccentBrush") };
        var cancel = new Button { Content = "Annulla", IsCancel = true, Width = 80 };
        ok.Click += (_, _) => DialogResult = true;
        var bar = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 14, 0, 0), Children = { ok, cancel } };
        Grid.SetRow(bar, r); Grid.SetColumnSpan(bar, 2); grid.Children.Add(bar);
        Content = grid;
        Loaded += (_, _) => { if (_inputs.Count > 0) { _inputs[0].Focus(); if (_inputs[0] is TextBox t) t.SelectAll(); } };
    }

    public string[] Values => _inputs.Select(i => i is TextBox t ? t.Text : (string)((ComboBox)i).SelectedItem).ToArray();

    /// <summary>Tutti i campi come numeri (virgola o punto); null se uno non è valido.</summary>
    public double[]? Numbers()
    {
        var res = new List<double>();
        foreach (var v in Values)
        {
            if (!double.TryParse(v.Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var d) || double.IsNaN(d) || Math.Abs(d) > 1e6) return null;
            res.Add(d);
        }
        return res.ToArray();
    }
}

/// <summary>Trova e sostituisci (non modale).</summary>
public sealed class FindDialog : Window
{
    readonly EditorSession _s;
    readonly Action _changed;
    readonly TextBox _find = new() { Height = 26, Margin = new Thickness(0, 4, 0, 4), VerticalContentAlignment = VerticalAlignment.Center };
    readonly TextBox _repl = new() { Height = 26, Margin = new Thickness(0, 4, 0, 4), VerticalContentAlignment = VerticalAlignment.Center };
    readonly CheckBox _case = new() { Content = "Maiuscole/minuscole", Margin = new Thickness(0, 6, 0, 6) };
    readonly TextBlock _msg = new() { Margin = new Thickness(0, 6, 0, 0), Foreground = (Brush)Application.Current.FindResource("MutedBrush") };
    readonly UIElement[] _replaceOnly;

    public FindDialog(Window owner, EditorSession s, Action changed)
    {
        _s = s; _changed = changed;
        Owner = owner; Title = "Trova e sostituisci"; SizeToContent = SizeToContent.WidthAndHeight; ResizeMode = ResizeMode.NoResize; ShowInTaskbar = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; Topmost = false;
        var g = new Grid { Margin = new Thickness(16), Width = 360 };
        g.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); g.ColumnDefinitions.Add(new ColumnDefinition());
        for (int i = 0; i < 5; i++) g.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        void Add(UIElement e, int row, int col) { Grid.SetRow(e, row); Grid.SetColumn(e, col); g.Children.Add(e); }
        Add(new TextBlock { Text = "Trova", Margin = new Thickness(0, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center }, 0, 0); Add(_find, 0, 1);
        var rl = new TextBlock { Text = "Sostituisci con", Margin = new Thickness(0, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center };
        Add(rl, 1, 0); Add(_repl, 1, 1);
        Grid.SetColumnSpan(_case, 2); Add(_case, 2, 0);
        var next = new Button { Content = "Trova successivo", IsDefault = true, Padding = new Thickness(10, 4, 10, 4) }; next.Click += (_, _) => FindNext();
        var rep = new Button { Content = "Sostituisci", Padding = new Thickness(10, 4, 10, 4) }; rep.Click += (_, _) => ReplaceOne();
        var all = new Button { Content = "Sostituisci tutto", Padding = new Thickness(10, 4, 10, 4) }; all.Click += (_, _) => ReplaceAll();
        var bar = new StackPanel { Orientation = Orientation.Horizontal, Children = { next, rep, all } };
        Grid.SetColumnSpan(bar, 2); Add(bar, 3, 0);
        Grid.SetColumnSpan(_msg, 2); Add(_msg, 4, 0);
        _replaceOnly = [rl, _repl, rep, all];
        Content = g;
        Loaded += (_, _) => { _find.Focus(); _find.SelectAll(); };
    }

    public void SetMode(bool replace)
    {
        foreach (var e in _replaceOnly) e.Visibility = replace ? Visibility.Visible : Visibility.Collapsed;
        if (_s.HasSelection && !_s.SelectedText().Contains('\n')) _find.Text = _s.SelectedText();
        _find.Focus(); _find.SelectAll();
    }

    void FindNext()
    {
        if (_find.Text.Length == 0) return;
        _msg.Text = _s.FindNext(_find.Text, _case.IsChecked == true) ? "" : "Nessun risultato.";
        _changed();
    }

    void ReplaceOne()
    {
        if (_find.Text.Length == 0) return;
        if (_s.HasSelection && string.Equals(_s.SelectedText(), _find.Text, _case.IsChecked == true ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase)) _s.Type(_repl.Text);
        FindNext();
    }

    void ReplaceAll()
    {
        if (_find.Text.Length == 0) return;
        int n = _s.ReplaceAll(_find.Text, _repl.Text, _case.IsChecked == true);
        _msg.Text = n == 0 ? "Nessun risultato." : $"{n} sostituzioni.";
        _changed();
    }
}

/// <summary>Stampa: ogni pagina è il disegno del layout, identico a quello a schermo.</summary>
public sealed class LayoutPaginator(LayoutResult layout) : DocumentPaginator
{
    const double K = 96.0 / 72.0;
    public override bool IsPageCountValid => true;
    public override int PageCount => layout.Pages.Count;
    public override Size PageSize { get => layout.Pages.Count > 0 ? new Size(layout.Pages[0].Width * K, layout.Pages[0].Height * K) : new Size(793, 1122); set { } }
    public override IDocumentPaginatorSource? Source => null;

    public override DocumentPage GetPage(int pageNumber)
    {
        var pg = layout.Pages[pageNumber];
        var v = new DrawingVisual();
        using (var dc = v.RenderOpen())
        {
            dc.PushTransform(new ScaleTransform(K, K));
            PageRenderer.Draw(dc, pg, 1.0, white: false);
            dc.Pop();
        }
        var size = new Size(pg.Width * K, pg.Height * K);
        return new DocumentPage(v, size, new Rect(size), new Rect(size));
    }
}
