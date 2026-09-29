using System;
using System.Runtime.InteropServices;
using Avalonia.Controls;
using DynamicGeometry;

namespace LiveGeometry.Desktop;

/// <summary>
/// The window's title bar in the theme's shade: Windows paints it light unless told, per
/// window, that the app is dark (the DWM attribute for immersive dark mode), and it is
/// told again whenever the theme changes.
/// </summary>
public static class WindowFrameTheme
{
    const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    public static void Attach(Window window)
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        window.Opened += (s, e) => Apply(window);
        AppTheme.CurrentChanged += () => Apply(window);
    }

    static void Apply(Window window)
    {
        var handle = window.TryGetPlatformHandle()?.Handle ?? IntPtr.Zero;
        if (handle == IntPtr.Zero)
        {
            return;
        }

        int dark = AppTheme.Current == AppTheme.Dark ? 1 : 0;
        DwmSetWindowAttribute(handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref dark, sizeof(int));
    }

    [DllImport("dwmapi.dll")]
    static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);
}
