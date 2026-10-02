using System;
using System.Globalization;
using Avalonia;
using Avalonia.Controls;

namespace LiveGeometry.Desktop;

/// <summary>
/// The main window comes back where it was closed on a Mac (and Linux): position, size,
/// maximized or not, through Avalonia itself - <see cref="WindowPlacementPersistence"/> is
/// the Windows way. Avalonia has no restored rectangle of a maximized window, so the bounds
/// are the last ones the window had while it was neither maximized nor full screen.
/// Kept in the <see cref="SettingsStore"/> under a key of its own.
/// </summary>
public static class WindowBoundsPersistence
{
    const string SettingKey = "WindowBounds";

    /// <summary>
    /// Call before the window is shown. Without saved bounds the window is left alone.
    /// </summary>
    public static void Attach(Window window)
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        try
        {
            Restore(window);
        }
        catch (Exception ex)
        {
            // never a reason not to start
            Console.WriteLine("Window bounds: " + ex.Message);
        }

        var bounds = new NormalBounds();
        window.PositionChanged += (s, e) => bounds.Changed(window);
        window.PropertyChanged += (s, e) =>
        {
            if (e.Property == Window.ClientSizeProperty || e.Property == Window.WindowStateProperty)
            {
                bounds.Changed(window);
            }
        };
        window.Closing += (s, e) => Save(window, bounds);
    }

    static void Restore(Window window)
    {
        var parts = (SettingsStore.Current.Get(SettingKey) ?? "").Split(',');
        if (parts.Length != 5
            || !int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int x)
            || !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int y)
            || !double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double width)
            || !double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out double height))
        {
            return;
        }

        // a sane size, whatever happened to the file
        if (width < 200 || height < 150)
        {
            return;
        }

        window.Width = width;
        window.Height = height;
        window.WindowState = parts[4] == nameof(WindowState.Maximized) ? WindowState.Maximized : WindowState.Normal;

        // only onto a screen that is still there: the window's top left, a little inside
        var corner = new PixelPoint(x + 40, y + 40);
        foreach (var screen in window.Screens.All)
        {
            if (screen.WorkingArea.Contains(corner))
            {
                window.WindowStartupLocation = WindowStartupLocation.Manual;
                window.Position = new PixelPoint(x, y);
                break;
            }
        }
    }

    static void Save(Window window, NormalBounds bounds)
    {
        try
        {
            if (window.WindowState == WindowState.Normal)
            {
                bounds.Settle(window.Position, window.ClientSize);
            }
            else
            {
                bounds.Changed(window);
            }

            if (bounds.Position == null)
            {
                // never normal since it opened maximized: nothing to come back to but that
                return;
            }

            var p = bounds.Position.Value;
            var s = bounds.Size.Value;
            var state = window.WindowState == WindowState.Maximized ? WindowState.Maximized : WindowState.Normal;
            SettingsStore.Current.Set(SettingKey, FormattableString.Invariant($"{p.X},{p.Y},{s.Width},{s.Height},{state}"));
        }
        catch (Exception ex)
        {
            Console.WriteLine("Window bounds: " + ex.Message);
        }
    }

    /// <summary>
    /// The bounds of the window while it is normal. A Mac animates a zoom: the window grows
    /// through a few sizes before it says it is maximized, and the last of those was taken for
    /// its normal size (1910 x 930 of a 1920 x 960 screen). So bounds count once they have held
    /// for a moment, or when the window closes with them.
    /// </summary>
    class NormalBounds
    {
        static readonly TimeSpan HoldTime = TimeSpan.FromSeconds(1);

        PixelPoint pendingPosition;
        Size pendingSize;
        DateTime pendingSince;
        bool isPending;

        public PixelPoint? Position { get; private set; }

        public Size? Size { get; private set; }

        /// <summary>The window moved, was resized or maximized.</summary>
        public void Changed(Window window)
        {
            // what came before has held long enough, or it was the way into this
            if (isPending && DateTime.UtcNow - pendingSince >= HoldTime)
            {
                Settle(pendingPosition, pendingSize);
            }

            isPending = window.WindowState == WindowState.Normal;
            pendingPosition = window.Position;
            pendingSize = window.ClientSize;
            pendingSince = DateTime.UtcNow;
        }

        public void Settle(PixelPoint position, Size size)
        {
            Position = position;
            Size = size;
        }
    }
}
