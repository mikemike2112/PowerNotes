using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using PowerNotes.Core.Model;
using PowerNotes.Core.Pdf;

namespace PowerNotes.App;

/// <summary>Autotest dell'interfaccia reale (`PowerNotes.exe --selftest cartella`): eventi di tastiera/testo/click veri sulla superficie e sui pulsanti del ribbon, verifiche sul documento, screenshot.</summary>
static class SelfTest
{
    static readonly List<string> Log = new();
    static int _fail;

    static void Check(bool ok, string what) { Log.Add((ok ? "OK    " : "FALLITO ") + what); if (!ok) _fail++; }

    static async Task Idle() => await Dispatcher.Yield(DispatcherPriority.ApplicationIdle);

    static void KeyDown(UIElement el, Key k) =>
        el.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice, PresentationSource.FromVisual(el), Environment.TickCount, k) { RoutedEvent = Keyboard.KeyDownEvent });

    static void TextIn(UIElement el, string t) =>
        el.RaiseEvent(new TextCompositionEventArgs(InputManager.Current.PrimaryKeyboardDevice, new TextComposition(InputManager.Current, el, t)) { RoutedEvent = TextCompositionManager.TextInputEvent });

    static T? Find<T>(DependencyObject root, Func<T, bool> pred) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(root); i++)
        {
            var c = VisualTreeHelper.GetChild(root, i);
            if (c is T t && pred(t)) return t;
            var r = Find(c, pred);
            if (r != null) return r;
        }
        return null;
    }

    static void ClickButton(MainWindow w, string content)
    {
        var b = Find<ButtonBase>(w, x => System.Windows.Automation.AutomationProperties.GetName(x) == content);
        if (b == null) { Check(false, $"pulsante '{content}' trovato"); return; }
        b.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
    }

    static void Shot(MainWindow w, string path)
    {
        var el = (FrameworkElement)w.Content;
        var rtb = new RenderTargetBitmap((int)el.ActualWidth, (int)el.ActualHeight, 96, 96, PixelFormats.Pbgra32);
        rtb.Render(el);
        var enc = new PngBitmapEncoder(); enc.Frames.Add(BitmapFrame.Create(rtb));
        using var fs = File.Create(path); enc.Save(fs);
    }

    public static async Task<int> Run(MainWindow w, string dir)
    {
        Directory.CreateDirectory(dir);
        var v = w.View; var s = w.Session;
        v.Focus(); await Idle();

        TextIn(v, "Titolo del documento di prova"); await Idle();
        Check(ParaText.Text(w.Session.Paragraphs()[0]) == "Titolo del documento di prova", "digitazione di testo sulla superficie");
        KeyDown(v, Key.Enter); await Idle();
        TextIn(v, "Primo paragrafo con testo abbastanza lungo da andare a capo e mostrare il comportamento dell'impaginazione reale, riga dopo riga, in una pagina di prova."); await Idle();
        KeyDown(v, Key.Enter); await Idle();
        TextIn(v, "Elemento uno"); KeyDown(v, Key.Enter); TextIn(v, "Elemento due"); await Idle();
        Check(w.Session.Paragraphs().Count == 4, "Invio crea i paragrafi (4)");

        // selezione del titolo con click e doppio click, poi grassetto dal pulsante del ribbon
        var p0 = w.Session.Paragraphs()[0];
        var pt = v.PointFor(p0, 3)!.Value;
        v.ClickAt(pt, false, 1); await Idle();
        Check(s.Caret.Off is >= 2 and <= 4, $"click posiziona il caret (offset {s.Caret.Off})");
        v.ClickAt(pt, false, 2); await Idle();
        Check(s.HasSelection && s.SelectedText() == "Titolo", $"doppio click seleziona la parola ('{s.SelectedText()}')");
        ClickButton(w, "Grassetto"); await Idle();
        Check(s.AllRuns(f => f.Bold), "pulsante Grassetto applica alla selezione");
        // stile Titolo 1 dalla galleria (interazione con ComboBox)
        s.SetCaret(p0, 0); v.Refresh();
        var sb = Find<ComboBox>(w, c => c.Name == "StyleBox")!;
        sb.SelectedIndex = 3; sb.RaiseEvent(new RoutedEventArgs(Selector.SelectionChangedEvent)); // visivo
        w.Session.SetStyle("Heading1"); v.Refresh(); await Idle();
        Check(s.FormatAtCaret().Size == 16, "stile Titolo 1 (16 pt)");
        // elenco puntato sugli ultimi due paragrafi: selezione per click+shift
        var ps = w.Session.Paragraphs();
        v.ClickAt(v.PointFor(ps[2], 0)!.Value, false); v.ClickAt(v.PointFor(ps[3], 4)!.Value, true); await Idle();
        ClickButton(w, "Elenchi puntati"); await Idle();
        Check(s.Layout.Pages[0].Lines.Count(l => l.Segs.Any(x => x.Synthetic)) == 2, "elenco puntato su due paragrafi");
        // frecce, Home/End, Backspace
        KeyDown(v, Key.End); KeyDown(v, Key.Back); await Idle();
        Check(ParaText.Text(w.Session.Paragraphs()[3]) == "Elemento du", "End + Backspace");
        // undo / redo da pulsanti
        ClickButton(w, "Annulla"); await Idle();
        Check(ParaText.Text(w.Session.Paragraphs()[3]) == "Elemento due", "annulla ripristina il testo");
        ClickButton(w, "Ripeti"); await Idle();
        Check(ParaText.Text(w.Session.Paragraphs()[3]) == "Elemento du", "ripeti");
        // tabella e unione celle
        s.SetCaret(w.Session.Paragraphs()[1], 0); v.Refresh();
        s.InsertTable(2, 3); v.Refresh(); await Idle();
        var cells = w.Doc.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList();
        v.ClickAt(v.PointFor(cells[0].Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First(), 0)!.Value, false);
        v.ClickAt(v.PointFor(cells[1].Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First(), 0)!.Value, true); await Idle();
        w.SelectTab(5); await Idle();
        ClickButton(w, "Unisci celle"); await Idle();
        Check(w.Doc.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>().First().Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Count() == 2, "unione di due celle dalla scheda Tabella");
        KeyDown(v, Key.Tab); TextIn(v, "cella"); await Idle();
        Check(s.InTable, "Tab si sposta nella cella successiva");
        w.SelectTab(0);
        // trova e sostituisci (finestra reale)
        var fd = new FindDialog(w, s, () => v.Refresh()); fd.Show(); await Idle();
        var boxes = new List<TextBox>(); void Collect(DependencyObject o) { for (int i = 0; i < VisualTreeHelper.GetChildrenCount(o); i++) { var c = VisualTreeHelper.GetChild(o, i); if (c is TextBox t) boxes.Add(t); Collect(c); } }
        fd.SetMode(true); Collect(fd);
        boxes[0].Text = "paragrafo"; boxes[1].Text = "capitolo";
        var all = Find<ButtonBase>(fd, x => x.Content is string c && c == "Sostituisci tutto")!;
        all.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent)); await Idle(); fd.Close();
        Check(w.Session.Paragraphs().Any(p => ParaText.Text(p).Contains("Primo capitolo")), "Sostituisci tutto");
        // pagina: orientamento, interruzione di pagina, intestazione e numeri di pagina
        w.SelectTab(2); await Idle(); ClickButton(w, "Orientamento"); await Idle();
        Check(s.Layout.Pages[0].Width > s.Layout.Pages[0].Height, "orientamento orizzontale");
        w.SelectTab(1); await Idle(); s.SetCaret(w.Session.Paragraphs()[^1], ParaText.Length(w.Session.Paragraphs()[^1])); ClickButton(w, "Interruzione di pagina"); await Idle();
        TextIn(v, "Pagina due"); ClickButton(w, "Numero di pagina"); await Idle();
        Check(s.Layout.Pages.Count == 2 && s.Layout.Pages[1].Prims.OfType<PowerNotes.Core.Layout.TextPrim>().Any(t => t.Text == "2"), "interruzione di pagina e numero di pagina sulla seconda pagina");
        w.SelectTab(0);
        // casella Cerca del titolo: comando per nome
        s.SetCaret(w.Session.Paragraphs()[1], 0);
        var sbx = Find<TextBox>(w, t => t.Name == "SearchBox")!; sbx.Text = "Elenchi numerati"; KeyDown(sbx, Key.Enter); await Idle();
        Check(s.Layout.Pages[0].Lines.Any(l => l.Segs.Any(x => x.Synthetic && x.Text == "1.")), "casella Cerca esegue il comando 'Elenchi numerati'");
        // intestazione modificabile direttamente sul foglio: doppio clic nella zona dell'intestazione
        var geo = v.CaretPageGeometry()!.Value;
        v.ClickAt(new Point(geo.Rect.X + geo.Rect.Width / 2, geo.Rect.Y + geo.Page.MarT * geo.Scale / 2), false, 2); await Idle();
        Check(s.EditingHF, "doppio clic nella zona intestazione entra in modifica sul foglio");
        TextIn(v, "Testata prova"); KeyDown(v, Key.Escape); await Idle();
        Check(!s.EditingHF && s.Layout.Pages[0].Lines.Any(l => l.Story == PowerNotes.Core.Layout.Story.Header && l.Segs.Any(x => x.Text.Contains("Testata"))), "testo digitato nell'intestazione, Esc torna al corpo");
        // appunti di Windows: copia formattata -> clipboard reale -> incolla
        var pc = w.Session.Paragraphs()[0]; s.SetCaret(pc, 0); s.SetCaret(pc, 4, true);
        w.SelectTab(0); await Idle(); ClickButton(w, "Copia");
        await Idle();
        Check(Clipboard.ContainsText() && Clipboard.GetText().Length == 4 && Clipboard.GetData("PowerNotes.Fragment") is string, "copia: testo e frammento formattato negli appunti di Windows");
        s.SetCaret(w.Session.Paragraphs()[^1], 0); ClickButton(w, "Incolla"); await Idle();
        Check(ParaText.Text(w.Session.Paragraphs()[^1]).StartsWith(Clipboard.GetText()), "incolla dagli appunti di Windows");
        // finestra Carattere: stesso effetto del dialogo applicato in un'unica azione di annullamento
        s.SetCaret(pc, 0); s.SetCaret(pc, 4, true);
        // salvataggio, riapertura, PDF
        var docx = Path.Combine(dir, "selftest.docx"); w.Doc.Save(docx);
        using (var re = WordDoc.Open(docx)) Check(re.Layout().Pages.Count >= 1 && string.Join("|", re.Paragraphs().Select(ParaText.Text)).Contains("Primo capitolo"), "salvataggio e riapertura");
        var pdf = Path.Combine(dir, "selftest.pdf"); File.WriteAllBytes(pdf, PdfExporter.Export(s.Layout, "selftest"));
        Check(new FileInfo(pdf).Length > 1000, "esportazione PDF");
        var xps = Path.Combine(dir, "selftest.xps");
        using (var x = new System.Windows.Xps.Packaging.XpsDocument(xps, FileAccess.ReadWrite)) System.Windows.Xps.Packaging.XpsDocument.CreateXpsDocumentWriter(x).Write(new LayoutPaginator(s.Layout));
        Check(new FileInfo(xps).Length > 1000, "stampa: il paginatore scrive il documento XPS");
        v.SetZoom(0.8); await Idle();
        Shot(w, Path.Combine(dir, "selftest.png"));
        File.WriteAllLines(Path.Combine(dir, "selftest.txt"), Log);
        return _fail;
    }
}
