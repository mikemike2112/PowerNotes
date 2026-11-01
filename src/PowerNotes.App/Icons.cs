using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Shapes;

namespace PowerNotes.App;

/// <summary>Set di icone proprio (tratti vettoriali su griglia 24x24). Nessuna risorsa grafica di terzi.</summary>
static class Icons
{
    static readonly Dictionary<string, string> D = new()
    {
        ["save"] = "M4,4 H17 L20,7 V20 H4 Z M8,4 V9 H15 V4 M7,20 V14 H17 V20",
        ["undo"] = "M9,6 L4,11 L9,16 M4,11 H14 A6,6 0 0 1 14,23 H11",
        ["redo"] = "M15,6 L20,11 L15,16 M20,11 H10 A6,6 0 0 0 10,23 H13",
        ["paste"] = "M7,5 H17 V21 H7 Z M10,3 H14 V7 H10 Z M10,12 H14 M10,16 H14",
        ["cut"] = "M7,3 L16,17 M17,3 L8,17 M7,20 A2,2 0 1 0 7.01,20 M17,20 A2,2 0 1 0 17.01,20",
        ["copy"] = "M8,8 H19 V21 H8 Z M5,16 V3 H15",
        ["brush"] = "M5,19 L9,15 M9,15 L18,6 L21,9 L12,18 Z",
        ["bullets"] = "M10,6 H21 M10,12 H21 M10,18 H21 M4,6 H6 M4,12 H6 M4,18 H6",
        ["numbers"] = "M10,6 H21 M10,12 H21 M10,18 H21 M4,4 V8 M3,13 H6 L3,17 H6",
        ["multilevel"] = "M4,5 H8 M8,10 H12 M12,15 H16 M10,20 H21 M14,5 H21 M18,10 H21",
        ["outdent"] = "M9,6 H21 M13,11 H21 M13,16 H21 M9,21 H21 M7,9 L3,13.5 L7,18",
        ["indent"] = "M9,6 H21 M13,11 H21 M13,16 H21 M9,21 H21 M3,9 L7,13.5 L3,18",
        ["sort"] = "M7,4 V20 M3,16 L7,20 L11,16 M15,6 H21 M15,12 H19 M15,18 H17",
        ["pilcrow"] = "M16,4 V20 M12,4 V20 M16,4 H9 A4.5,4.5 0 0 0 9,13 H12",
        ["alignL"] = "M4,6 H20 M4,11 H14 M4,16 H20 M4,21 H12",
        ["alignC"] = "M4,6 H20 M7,11 H17 M4,16 H20 M8,21 H16",
        ["alignR"] = "M4,6 H20 M10,11 H20 M4,16 H20 M12,21 H20",
        ["alignJ"] = "M4,6 H20 M4,11 H20 M4,16 H20 M4,21 H20",
        ["linespace"] = "M9,6 H21 M9,12 H21 M9,18 H21 M5,4 V20 M2,7 L5,4 L8,7 M2,17 L5,20 L8,17",
        ["shading"] = "M5,13 L12,6 L19,13 L12,20 Z M20,15 C20,18 22,19 22,21 A2,2 0 0 1 18,21 C18,19 20,18 20,15",
        ["borders"] = "M4,4 H20 V20 H4 Z M4,12 H20 M12,4 V20",
        ["find"] = "M10,4 A6,6 0 1 0 10.01,4 M14.5,14.5 L21,21",
        ["replace"] = "M4,8 H15 L12,5 M20,16 H9 L12,19",
        ["select"] = "M6,3 V19 L10,15 L13,22 L16,21 L13,14 H19 Z",
        ["table"] = "M3,4 H21 V20 H3 Z M3,10 H21 M3,15 H21 M9,4 V20 M15,4 V20",
        ["image"] = "M3,5 H21 V19 H3 Z M3,16 L9,10 L14,15 L17,12 L21,16 M16,9 A1.5,1.5 0 1 0 16.01,9",
        ["pagebreak"] = "M7,3 H17 V9 M7,3 V9 M3,12 H6 M9,12 H15 M18,12 H21 M7,15 V21 H17 V15",
        ["blankpage"] = "M6,3 H14 L19,8 V21 H6 Z M14,3 V8 H19",
        ["header"] = "M5,3 H19 V21 H5 Z M5,8 H19 M8,5.5 H14",
        ["footer"] = "M5,3 H19 V21 H5 Z M5,16 H19 M8,18.5 H14",
        ["pagenum"] = "M5,3 H19 V21 H5 Z M9,15 H15 M12,12 V18",
        ["textbox"] = "M3,5 H21 V19 H3 Z M9,16 L12,8 L15,16 M10,13.5 H14",
        ["margins"] = "M4,3 H20 V21 H4 Z M7,6 H17 V18 H7 Z",
        ["orient"] = "M7,3 H17 V15 H7 Z M3,19 H21 M18,16 L21,19 L18,22",
        ["size"] = "M6,3 H18 V21 H6 Z M3,3 V21 M21,3 V21",
        ["columns"] = "M4,4 H20 V20 H4 Z M12,4 V20 M6.5,8 H9.5 M6.5,12 H9.5 M14.5,8 H17.5 M14.5,12 H17.5",
        ["breaks"] = "M4,6 H20 M4,10 H14 M3,15 H7 M10,15 H14 M17,15 H21 M4,20 H20",
        ["zoom"] = "M10,4 A6,6 0 1 0 10.01,4 M14.5,14.5 L21,21 M8,10 H12 M10,8 V12",
        ["zoom100"] = "M6,3 H15 L19,7 V21 H6 Z M9,14 H16 M9,17 H13",
        ["onepage"] = "M7,3 H17 V21 H7 Z",
        ["pagewidth"] = "M5,6 H19 V18 H5 Z M2,12 H5 M19,12 H22 M3.5,10 V14 M20.5,10 V14",
        ["read"] = "M3,5 H11 V19 H3 Z M13,5 H21 V19 H13 Z M5,9 H9 M15,9 H19",
        ["print"] = "M7,3 H17 V8 H7 Z M5,8 H19 V17 H5 Z M8,13 H16 V21 H8 Z",
        ["web"] = "M12,3 A9,9 0 1 0 12.01,3 M3,12 H21 M12,3 C7,8 7,16 12,21 M12,3 C17,8 17,16 12,21",
        ["outline"] = "M4,6 H20 M8,11 H20 M12,16 H20 M4,21 H20",
        ["rowabove"] = "M3,12 H21 V21 H3 Z M12,3 V9 M9,6 H15",
        ["rowbelow"] = "M3,3 H21 V12 H3 Z M12,15 V21 M9,18 H15",
        ["colleft"] = "M12,3 H21 V21 H12 Z M3,12 H9 M6,9 V15",
        ["colright"] = "M3,3 H12 V21 H3 Z M15,12 H21 M18,9 V15",
        ["delete"] = "M4,4 H16 V16 H4 Z M4,10 H16 M10,4 V16 M15,15 L21,21 M21,15 L15,21",
        ["merge"] = "M3,5 H21 V19 H3 Z M3,12 H8 M16,12 H21 M10,9 L7,12 L10,15 M14,9 L17,12 L14,15",
        ["split"] = "M3,5 H21 V19 H3 Z M12,5 V19 M9,9 L6,12 L9,15 M15,9 L18,12 L15,15",
        ["new"] = "M6,3 H14 L19,8 V21 H6 Z M14,3 V8 H19 M12,12 V18 M9,15 H15",
        ["open"] = "M3,6 H10 L12,8 H21 V19 H3 Z M3,11 H21",
        ["export"] = "M6,3 H14 L19,8 V21 H6 Z M14,3 V8 H19 M9,16 L12,19 L15,16 M12,12 V19",
        ["close"] = "M5,5 L19,19 M19,5 L5,19",
        ["gear"] = "M12,8 A4,4 0 1 0 12.01,8 M12,2 V5 M12,19 V22 M2,12 H5 M19,12 H22 M5,5 L7,7 M17,17 L19,19 M19,5 L17,7 M7,17 L5,19",
        ["moon"] = "M20,14 A8,8 0 1 1 10,4 A6.5,6.5 0 0 0 20,14 Z",
        ["dlg-info"] = "M12,3 A9,9 0 1 0 12.01,3 M12,11 V17 M12,7.5 V8",
        ["dlg-warn"] = "M12,3 L22,20 H2 Z M12,10 V15 M12,17.5 V18",
        ["dlg-error"] = "M12,3 A9,9 0 1 0 12.01,3 M8.5,8.5 L15.5,15.5 M15.5,8.5 L8.5,15.5",
        ["dlg-question"] = "M12,3 A9,9 0 1 0 12.01,3 M9.5,9.5 A2.5,2.5 0 1 1 12,12 V14 M12,16.5 V17.5",
        ["min"] = "M5,12 H19",
        ["max"] = "M6,6 H18 V18 H6 Z",
        ["generic"] = "M5,5 H19 V19 H5 Z",
        ["caret"] = "M6,9 L12,15 L18,9",
        ["search"] = "M10,4 A6,6 0 1 0 10.01,4 M14.5,14.5 L21,21",
    };

    static readonly Dictionary<string, Geometry> Cache = new();

    public static Geometry Geo(string key)
    {
        if (Cache.TryGetValue(key, out var g)) return g;
        g = Geometry.Parse(D.TryGetValue(key, out var d) ? d : D["generic"]); g.Freeze();
        return Cache[key] = g;
    }

    public static FrameworkElement Make(string key, double size, string brushKey = "TextBrush")
    {
        var p = new Path { Data = Geo(key), StrokeThickness = 1.4, StrokeLineJoin = PenLineJoin.Round, StrokeStartLineCap = PenLineCap.Round, StrokeEndLineCap = PenLineCap.Round, Width = 24, Height = 24, Fill = Brushes.Transparent };
        p.SetResourceReference(Shape.StrokeProperty, brushKey);
        return new Viewbox { Width = size, Height = size, Child = p, Stretch = Stretch.Uniform };
    }
}
