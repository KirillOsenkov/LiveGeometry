using System;
using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform;
using Avalonia.Themes.Fluent;
using Avalonia.Threading;
using DynamicGeometry;

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

        // the chrome's colors under each theme variant; the stored choice (or the system's)
        // before the first frame, so nothing flashes light first. The stored colors go on
        // before the dictionaries are registered: they go straight in then, where afterwards
        // they would wait for an idle moment (AppTheme.SetResource), after the first frame.
        AppSettings.Instance.Load();
        AppTheme.Register(this);

        // An exception that gets out of an event handler or a posted job would end the app
        // (the desktop window disappears, the browser page freezes) and take the drawing with
        // it. It has been shown already, at the throw (MainView.CurrentDomain_FirstChanceException):
        // here it is only kept from going any further.
        Dispatcher.UIThread.UnhandledException += (s, e) => e.Handled = true;

        // A Mac's menu bar names the app and has an application menu: without a name and a menu
        // of our own they said "Avalonia Application" and "About Avalonia". Avalonia reads both
        // after this, and puts Services, Hide and Quit after our items.
        Name = "Live Geometry";
        if (OperatingSystem.IsMacOS())
        {
            NativeMenu.SetMenu(this, CreateMacApplicationMenu());
        }
    }

    NativeMenu CreateMacApplicationMenu()
    {
        var gitHub = new NativeMenuItem("Live Geometry on GitHub") { ToolTip = BuildVersion.Full };
        gitHub.Click += (s, e) =>
        {
            var window = (ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;
            window?.Launcher.LaunchUriAsync(new Uri(MainView.RepositoryUrl));
        };

        return new NativeMenu() { gitHub };
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
            window.Closing += (s, e) => SettingsStore.RaiseLeaving();

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
