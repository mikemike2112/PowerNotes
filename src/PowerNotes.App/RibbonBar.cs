using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace PowerNotes.App;

/// <summary>Contenuto di una scheda del ribbon che si adatta alla larghezza: prima i gruppi si compattano (solo icone), poi quelli in più passano in un menu a discesa.</summary>
public sealed class RibbonBar : Grid
{
    public sealed class GroupInfo { public required FrameworkElement Root; public List<UIElement> Labels = new(); public bool Compact; public bool CanCompact => Labels.Count > 0; }

    readonly StackPanel _main = new() { Orientation = Orientation.Horizontal };
    readonly ToggleButton _more = new() { Content = "⋯", Padding = new Thickness(10, 4, 10, 4), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0, 0, 0), ToolTip = "Altri comandi", Visibility = Visibility.Collapsed };
    readonly Popup _popup = new() { StaysOpen = false, AllowsTransparency = false, Placement = PlacementMode.Bottom };
    readonly WrapPanel _overflowHost = new() { MaxWidth = 520, Margin = new Thickness(6) };
    readonly List<GroupInfo> _groups = new();
    bool _busy;

    /// <summary>Numero di gruppi compattati e di gruppi nel menu (per test e diagnosi).</summary>
    public int CompactCount => _groups.Count(g => g.Compact);
    public int OverflowCount => _overflowHost.Children.Count;

    public RibbonBar()
    {
        Children.Add(_main);
        _more.HorizontalAlignment = HorizontalAlignment.Right;
        Children.Add(_more);
        var border = new Border { BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Child = _overflowHost };
        border.SetResourceReference(Border.BackgroundProperty, "PanelBrush"); border.SetResourceReference(Border.BorderBrushProperty, "LineBrush");
        _popup.Child = border; _popup.PlacementTarget = _more; _popup.SetBinding(Popup.IsOpenProperty, new System.Windows.Data.Binding("IsChecked") { Source = _more, Mode = System.Windows.Data.BindingMode.TwoWay });
        System.Windows.Automation.AutomationProperties.SetName(_more, "Altri comandi");
        SizeChanged += (_, e) => { if (e.WidthChanged) Relayout(); };
        Loaded += (_, _) => Relayout();
    }

    public GroupInfo AddGroup(FrameworkElement root)
    {
        var g = new GroupInfo { Root = root };
        _groups.Add(g); _main.Children.Add(root);
        return g;
    }

    static void SetCompact(GroupInfo g, bool on)
    {
        g.Compact = on;
        foreach (var l in g.Labels) l.Visibility = on ? Visibility.Collapsed : Visibility.Visible;
    }

    double NaturalWidth()
    {
        _main.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
        return _main.DesiredSize.Width;
    }

    /// <summary>Ricalcola compattamento e menu per la larghezza attuale.</summary>
    public void Relayout()
    {
        if (_busy) return; _busy = true;
        try
        {
            double avail = ActualWidth;
            if (avail <= 0) return;
            // tutto al massimo
            foreach (var g in _overflowHost.Children.OfType<FrameworkElement>().ToList()) { _overflowHost.Children.Remove(g); }
            _main.Children.Clear();
            foreach (var g in _groups) { SetCompact(g, false); _main.Children.Add(g.Root); }
            _more.Visibility = Visibility.Collapsed; _more.IsChecked = false;
            if (NaturalWidth() <= avail) return;
            // 1) compatta da destra a sinistra
            for (int i = _groups.Count - 1; i >= 0 && NaturalWidth() > avail; i--) if (_groups[i].CanCompact) SetCompact(_groups[i], true);
            if (NaturalWidth() <= avail) return;
            // 2) i gruppi in eccesso (dall'ultimo) vanno nel menu "⋯"
            _more.Visibility = Visibility.Visible;
            double limit = avail - 56;
            for (int i = _groups.Count - 1; i > 0 && NaturalWidth() > limit; i--)
            {
                var g = _groups[i];
                _main.Children.Remove(g.Root); SetCompact(g, false);
                _overflowHost.Children.Insert(0, g.Root);
            }
        }
        finally { _busy = false; }
    }
}
