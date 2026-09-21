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

        App.MainWindowCreated = WindowPlacementPersistence.Attach;
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
