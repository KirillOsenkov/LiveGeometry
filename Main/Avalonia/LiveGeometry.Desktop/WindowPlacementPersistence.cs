using System;
using System.IO;
using System.Runtime.InteropServices;
using Avalonia.Controls;

namespace LiveGeometry.Desktop;

/// <summary>
/// The main window comes back where it was closed: position, size, maximized or not.
/// Windows only, through Get/SetWindowPlacement (the same approach as in Helix and
/// MSBuild Structured Log Viewer): the placement is about the *restored* rectangle even while
/// the window is maximized, it is in work area coordinates, and Windows itself pulls a window
/// back onto a screen if the monitor it was on is gone.
/// </summary>
public static class WindowPlacementPersistence
{
    const int SW_HIDE = 0;
    const int SW_SHOWNORMAL = 1;
    const int SW_SHOWMINIMIZED = 2;
    const int SW_SHOWMAXIMIZED = 3;

    public static string SettingsFile
    {
        get
        {
            var folder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(folder, "LiveGeometry", "MainWindowPosition.txt");
        }
    }

    /// <summary>
    /// Call before the window is shown. Without a saved placement the window is left alone.
    /// </summary>
    public static void Attach(Window window)
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        window.Closing += (s, e) => Save(window);

        try
        {
            Restore(window);
        }
        catch (Exception ex)
        {
            // never a reason not to start
            Console.WriteLine("Window placement: " + ex.Message);
        }
    }

    static void Restore(Window window)
    {
        if (!File.Exists(SettingsFile))
        {
            return;
        }

        var placement = WindowPlacement.Parse(File.ReadAllText(SettingsFile).Trim());
        var handle = GetHandle(window);
        if (placement == null || handle == IntPtr.Zero)
        {
            return;
        }

        // The rectangle now, invisibly; showing is Avalonia's job, in the state that was saved
        // (a window that was closed minimized comes back normal).
        bool maximized = placement.showCmd == SW_SHOWMAXIMIZED;
        placement.showCmd = SW_HIDE;
        if (SetWindowPlacement(handle, placement))
        {
            window.WindowStartupLocation = WindowStartupLocation.Manual;
            window.WindowState = maximized ? WindowState.Maximized : WindowState.Normal;
        }
    }

    static void Save(Window window)
    {
        try
        {
            var handle = GetHandle(window);
            if (handle == IntPtr.Zero)
            {
                Console.WriteLine("Window placement: no window handle to save from");
                return;
            }

            var placement = new WindowPlacement();
            if (!GetWindowPlacement(handle, placement))
            {
                Console.WriteLine("Window placement: GetWindowPlacement failed, error " + Marshal.GetLastWin32Error());
                return;
            }

            if (placement.showCmd == SW_SHOWMINIMIZED)
            {
                placement.showCmd = SW_SHOWNORMAL;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(SettingsFile));
            File.WriteAllText(SettingsFile, placement.ToString());
        }
        catch (Exception ex)
        {
            Console.WriteLine("Window placement: " + ex.Message);
        }
    }

    static IntPtr GetHandle(Window window)
    {
        var platformHandle = window.TryGetPlatformHandle();
        return platformHandle != null ? platformHandle.Handle : IntPtr.Zero;
    }

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool SetWindowPlacement(IntPtr hWnd, WindowPlacement placement);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool GetWindowPlacement(IntPtr hWnd, WindowPlacement placement);

    [StructLayout(LayoutKind.Sequential)]
    public class WindowPlacement
    {
        public int length = Marshal.SizeOf(typeof(WindowPlacement));
        public int flags;
        public int showCmd;
        public int minX;
        public int minY;
        public int maxX;
        public int maxY;
        public int left;
        public int top;
        public int right;
        public int bottom;

        public static WindowPlacement Parse(string text)
        {
            var parts = (text ?? "").Split(',');
            var numbers = new int[10];
            if (parts.Length != numbers.Length)
            {
                return null;
            }

            for (int i = 0; i < numbers.Length; i++)
            {
                if (!int.TryParse(parts[i], out numbers[i]))
                {
                    return null;
                }
            }

            // a sane size, whatever happened to the file
            if (numbers[8] < 200 || numbers[9] < 150)
            {
                return null;
            }

            return new WindowPlacement()
            {
                flags = numbers[0],
                showCmd = numbers[1],
                minX = numbers[2],
                minY = numbers[3],
                maxX = numbers[4],
                maxY = numbers[5],
                left = numbers[6],
                top = numbers[7],
                right = numbers[6] + numbers[8],
                bottom = numbers[7] + numbers[9]
            };
        }

        // flags, showCmd, minimized x y, maximized x y, left, top, width, height - as in Helix
        public override string ToString()
        {
            return $"{flags},{showCmd},{minX},{minY},{maxX},{maxY},{left},{top},{right - left},{bottom - top}";
        }
    }
}
