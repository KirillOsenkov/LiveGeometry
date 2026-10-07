using System;
using System.IO;
using System.Threading.Tasks;
using Avalonia;

namespace LiveGeometry.Desktop;

sealed class Program
{
    // Initialization code. Don't use any Avalonia, third-party APIs or any
    // SynchronizationContext-reliant code before AppMain is called: things aren't initialized
    // yet and stuff might break.
    [STAThread]
    public static void Main(string[] args)
    {
        App.UseInvariantCulture();
        SettingsStore.Current = new FileSettingsStore();

        // "LiveGeometry.Desktop.exe drawing.lgf", which is also what a file association runs
        if (args.Length > 0 && System.IO.File.Exists(args[0]))
        {
            MainView.StartupFile = System.IO.Path.GetFullPath(args[0]);
        }
        else if (args.Length >= 3 && args[0] == "--check")
        {
            // load every drawing under a folder and write a picture and a report of each
            MainView.CheckFolder = System.IO.Path.GetFullPath(args[1]);
            MainView.CheckOutputFolder = System.IO.Path.GetFullPath(args[2]);
            MainView.CheckDark = args.Length > 3 && args[3] == "--dark";
        }
        else if (args.Length == 2 && args[0] == "--modernize")
        {
            MainView.ModernizeFolder = System.IO.Path.GetFullPath(args[1]);
        }
        else if (args.Length >= 2 && args[0] == "--rewrite")
        {
            // every drawing of a folder loaded and saved again, in today's format; with
            // --letters, its figures named Circle1 renamed as a new drawing names them (c)
            MainView.RewriteFolder = System.IO.Path.GetFullPath(args[1]);
            MainView.RewriteGivesLetters = args.Length > 2 && args[2] == "--letters";
        }
        else if (args.Length == 2 && args[0] == "--recaption")
        {
            MainView.RecaptionFolder = System.IO.Path.GetFullPath(args[1]);
        }
        else if (args.Length == 2 && args[0] == "--space-labels")
        {
            MainView.SpaceLabelsFolder = System.IO.Path.GetFullPath(args[1]);
        }
        else if (args.Length == 2 && args[0] == "--bench-expressions")
        {
            // the expression strategies measured over the gallery, the table written to the file
            MainView.BenchmarkOutput = System.IO.Path.GetFullPath(args[1]);
        }
        else if (args.Length == 1 && args[0] == "--arrange")
        {
            // reorder the gallery by dragging its tiles; the catalog's source is rewritten
            MainView.ArrangeGallery = true;
        }
        else if (args.Length == 2 && args[0] == "--gallery")
        {
            // straight to a drawing of the gallery, as the browser's /gallery/<slug> would
            MainView.StartupPath = "/gallery/" + args[1];
        }

        DynamicGeometry.EmojiFont.Open = () => Task.FromResult<Stream>(
            File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Fonts", "Twemoji.Mozilla.ttf")));
        App.MainWindowCreated = window =>
        {
            WindowPlacementPersistence.Attach(window);
            WindowBoundsPersistence.Attach(window);
            WindowFrameTheme.Attach(window);
            MacDockIcon.Apply();
        };
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
