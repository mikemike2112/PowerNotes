using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace PowerNotes.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += (_, ex) =>
        {
            ThemedDialog.Show("Errore imprevisto: " + ex.Exception.Message + "\n\nSalva il lavoro e riavvia PowerNotes.", "PowerNotes", MessageBoxButton.OK, MessageBoxImage.Error);
            ex.Handled = true;
        };
        var w = new MainWindow();
        string? dlgShot = null; string? shot = null; double zoom = 1; bool light = false; int tab = 0; int sel = 0;
        string? path = null; string? selftest = null;
        for (int i = 0; i < e.Args.Length; i++)
        {
            switch (e.Args[i])
            {
                case "--shot": shot = e.Args[++i]; break;
                case "--zoom": zoom = double.Parse(e.Args[++i], System.Globalization.CultureInfo.InvariantCulture); break;
                case "--light": light = true; break;
                case "--tab": tab = int.Parse(e.Args[++i]); break;
                case "--select": sel = int.Parse(e.Args[++i]); break;
                case "--dialog-shot": dlgShot = e.Args[++i]; break;
                case "--selftest": selftest = e.Args[++i]; break;
                default: path ??= e.Args[i]; break;
            }
        }
        if (path != null) w.OpenPath(path);
        if (light) w.SetTheme(true);
        w.View.SetZoom(zoom);
        w.SelectTab(tab);
        if (shot == null && selftest == null) w.WindowState = WindowState.Maximized;
        w.Show();
        if (dlgShot != null)
        {
            // schermata della richiesta di salvataggio: PrintWindow della sola finestra del dialogo, nessun input
            w.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, async () =>
            {
                var d = ThemedDialog.PreviewSave(w, "Relazione.docx");
                await Dispatcher.Yield(DispatcherPriority.ApplicationIdle); await Task.Delay(400);
                DialogShot.Capture(d, dlgShot);
                d.Close(); Shutdown();
            });
        }
        if (selftest != null)
        {
            w.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, async () =>
            {
                int fails;
                try { fails = await SelfTest.Run(w, selftest); }
                catch (Exception ex) { File.WriteAllText(Path.Combine(selftest, "selftest-errore.txt"), ex.ToString()); fails = 99; }
                Shutdown(fails);
            });
        }
        if (shot != null)
        {
            // modalità schermata per README/verifica: disegna la finestra reale su PNG ed esce
            w.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, () =>
            {
                if (sel > 0) w.ApplyDemoSelection(sel);
                w.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, () =>
                {
                    var el = (FrameworkElement)w.Content;
                    var rtb = new RenderTargetBitmap((int)el.ActualWidth, (int)el.ActualHeight, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
                    rtb.Render(el);
                    var enc = new PngBitmapEncoder(); enc.Frames.Add(BitmapFrame.Create(rtb));
                    using (var fs = File.Create(shot)) enc.Save(fs);
                    Shutdown();
                });
            });
        }
    }
}
