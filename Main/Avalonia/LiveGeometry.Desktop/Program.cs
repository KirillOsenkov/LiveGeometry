using System;
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
        // "LiveGeometry.Desktop.exe drawing.lgf", which is also what a file association runs
        if (args.Length > 0 && System.IO.File.Exists(args[0]))
        {
            MainView.StartupFile = System.IO.Path.GetFullPath(args[0]);
        }
        else if (args.Length == 3 && args[0] == "--check")
        {
            // load every drawing under a folder and write a picture and a report of each
            MainView.CheckFolder = System.IO.Path.GetFullPath(args[1]);
            MainView.CheckOutputFolder = System.IO.Path.GetFullPath(args[2]);
        }
        else if (args.Length == 2 && args[0] == "--modernize")
        {
            MainView.ModernizeFolder = System.IO.Path.GetFullPath(args[1]);
        }
        else if (args.Length == 2 && args[0] == "--recaption")
        {
            MainView.RecaptionFolder = System.IO.Path.GetFullPath(args[1]);
        }
        else if (args.Length == 2 && args[0] == "--space-labels")
        {
            MainView.SpaceLabelsFolder = System.IO.Path.GetFullPath(args[1]);
        }
        else if (args.Length == 2 && args[0] == "--gallery")
        {
            // straight to a drawing of the gallery, as the browser's /gallery/<slug> would
            MainView.StartupPath = "/gallery/" + args[1];
        }

        App.MainWindowCreated = WindowPlacementPersistence.Attach;
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
