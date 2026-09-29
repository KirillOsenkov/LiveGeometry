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

        // Windows 10 keeps painting the old shade until the frame is next redrawn (on the
        // next activation, otherwise): an activation cycle on the non-client area, and a
        // frame-changed notice, make it redraw now
        bool active = GetForegroundWindow() == handle;
        SendMessage(handle, WM_NCACTIVATE, active ? IntPtr.Zero : (IntPtr)1, IntPtr.Zero);
        SendMessage(handle, WM_NCACTIVATE, active ? (IntPtr)1 : IntPtr.Zero, IntPtr.Zero);
        SetWindowPos(handle, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    }

    const int WM_NCACTIVATE = 0x86;
    const uint SWP_NOSIZE = 0x0001;
    const uint SWP_NOMOVE = 0x0002;
    const uint SWP_NOZORDER = 0x0004;
    const uint SWP_NOACTIVATE = 0x0010;
    const uint SWP_FRAMECHANGED = 0x0020;

    [DllImport("dwmapi.dll")]
    static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    static extern IntPtr SendMessage(IntPtr hwnd, int message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int width, int height, uint flags);
}
