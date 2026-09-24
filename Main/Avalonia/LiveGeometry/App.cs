using System;
using System.Globalization;
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

    /// <summary>
    /// First thing in every head's Main. Numbers are written and read with a point
    /// everywhere - the expression language, the files, Avalonia's path markup, labels -
    /// and the UI is English, so the user's culture has nothing to say. With it in charge,
    /// a browser set to German formatted 17.5 as "17,5" into the app icon's path data and
    /// the app died at startup.
    /// </summary>
    public static void UseInvariantCulture()
    {
        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;
        CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.CurrentUICulture = CultureInfo.InvariantCulture;
    }

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

            AddressBar.Current.TitleChanged += title => window.Title = title;

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
