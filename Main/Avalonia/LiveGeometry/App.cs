using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform;
using Avalonia.Styling;
using Avalonia.Themes.Fluent;

namespace LiveGeometry;

public class App : Application
{
    /// <summary>
    /// For the head that has a window, to do what only it can: called before the window shows
    /// (the desktop head restores the window position here).
    /// </summary>
    public static Action<Window> MainWindowCreated { get; set; }

    public override void Initialize()
    {
        Styles.Add(new FluentTheme());
        RequestedThemeVariant = ThemeVariant.Light;
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var window = new Window
            {
                Title = "Live Geometry",
                Content = new MainView(),
                WindowState = WindowState.Maximized,
                Icon = new WindowIcon(AssetLoader.Open(new Uri("avares://LiveGeometry/Assets/DG.ico")))
            };

            if (MainWindowCreated != null)
            {
                MainWindowCreated(window);
            }

            desktop.MainWindow = window;
        }
        else if (ApplicationLifetime is ISingleViewApplicationLifetime singleView)
        {
            singleView.MainView = new MainView();
        }

        base.OnFrameworkInitializationCompleted();
    }
}
