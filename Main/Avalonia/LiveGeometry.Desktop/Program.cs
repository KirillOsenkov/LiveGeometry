using System;
using System.IO;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;

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
            // load every drawing under a folder and write a picture and a report of each;
            // "--dark" for the dark theme, "--size 390x645" for a window of that client size
            // (a phone's), which then leaves the window's saved place alone
            MainView.CheckFolder = System.IO.Path.GetFullPath(args[1]);
            MainView.CheckOutputFolder = System.IO.Path.GetFullPath(args[2]);
            for (int i = 3; i < args.Length; i++)
            {
                if (args[i] == "--dark")
                {
                    MainView.CheckDark = true;
                }
                else if (args[i] == "--size" && i + 1 < args.Length)
                {
                    checkSize = Size.Parse(args[++i].Replace('x', ','));
                }
            }
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

        LiveGeometry.EmojiFont.Open = () => Task.FromResult<Stream>(
            File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Fonts", "Twemoji.Mozilla.ttf")));

        // the JavaScript player, for Save as .html (copied beside the exe by the project)
        PlayerEmbed.OpenPlayerScript = () => File.ReadAllTextAsync(Path.Combine(AppContext.BaseDirectory, "player", "player.js"));
        App.MainWindowCreated = window =>
        {
            if (checkSize is Size size)
            {
                window.WindowState = WindowState.Normal;
                window.Width = size.Width;
                window.Height = size.Height;
            }
            else
            {
                WindowPlacementPersistence.Attach(window);
                WindowBoundsPersistence.Attach(window);
            }

            WindowFrameTheme.Attach(window);
            MacDockIcon.Apply();
        };
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    /// <summary>"--check ... --size": the window's client size, in layout pixels</summary>
    static Size? checkSize;

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
