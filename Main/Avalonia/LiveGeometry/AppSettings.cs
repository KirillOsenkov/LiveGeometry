using System;
using System.ComponentModel;
using DynamicGeometry;

namespace LiveGeometry;

/// <summary>
/// The settings page (the gear on the toolbar): what the app remembers between runs, read
/// from the <see cref="SettingsStore"/> at startup and written back as it is changed here.
/// Shown in the property grid like a figure, without undo.
/// </summary>
[PropertyGridName("Settings")]
[PropertyGridNoUndo]
public class AppSettings : INotifyPropertyChanged
{
    public static AppSettings Instance { get; } = new AppSettings();

    const string ThemeKey = "Theme";

    public event PropertyChangedEventHandler PropertyChanged;

    /// <summary>Something to show in the side panel: the theme's colors</summary>
    public event Action<object> ShowRequested;

    /// <summary>Before the first window: the stored choices take effect</summary>
    public void Load()
    {
        theme = SettingsStore.Current.Get(ThemeKey) ?? AppTheme.SystemChoice;
        if (theme != AppTheme.SystemChoice && AppTheme.ByName(theme) == null)
        {
            theme = AppTheme.SystemChoice;
        }

        AppTheme.Apply(theme);
    }

    string theme = AppTheme.SystemChoice;

    /// <summary>
    /// The name of the theme, or <see cref="AppTheme.SystemChoice"/> to follow the operating
    /// system (the browser)
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridPreferredEditor("ThemeChoice")]
    public string Theme
    {
        get => theme;
        set
        {
            if (theme == value)
            {
                return;
            }

            theme = value;
            AppTheme.Apply(value);
            SettingsStore.Current.Set(ThemeKey, value == AppTheme.SystemChoice ? null : value);
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Theme)));
        }
    }

    /// <summary>
    /// The sun/moon button: the other of light and dark. Landing on what the system asks for
    /// goes back to following the system, so the button undoes itself.
    /// </summary>
    public void ToggleDarkTheme()
    {
        var next = AppTheme.Current == AppTheme.Dark ? AppTheme.Light : AppTheme.Dark;
        Theme = next == AppTheme.System ? AppTheme.SystemChoice : next.Name;
    }

    /// <summary>The colors of the theme on screen, to tweak by eye</summary>
    [PropertyGridVisible]
    [PropertyGridName("Theme colors")]
    [PropertyGridIcon(PropertyGridIcon.Pencil)]
    public void EditThemeColors()
    {
        ShowRequested?.Invoke(AppTheme.Current);
    }
}
