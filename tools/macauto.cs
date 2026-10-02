#:property AllowUnsafeBlocks=true
#:property Nullable=disable
#:property PublishAot=false

// macauto - minimal macOS UI automation for the desktop app, the counterpart of winauto.cs.
//
//   dotnet tools/macauto.cs -- list
//   dotnet tools/macauto.cs -- shot   <target> <out.png> [--screen]
//   dotnet tools/macauto.cs -- click  <target> <x> <y> [left|right|double] [--shift] [--alt] [--cmd] [--ctrl]
//   dotnet tools/macauto.cs -- drag   <target> <x1> <y1> <x2> <y2> [steps] [--shift] [--alt] [--cmd] [--ctrl]  (held throughout)
//   dotnet tools/macauto.cs -- move   <target> <x> <y>
//   dotnet tools/macauto.cs -- wheel  <target> <x> <y> <notches>   (positive: up)
//   dotnet tools/macauto.cs -- key    <target> <key> [cmd] [ctrl] [shift] [alt]   e.g. "s cmd", "Escape", "Return"
//   dotnet tools/macauto.cs -- keys   <target> <text>             (each character as its key: tool letters)
//   dotnet tools/macauto.cs -- text   <target> <literal text>     (as Unicode, for text boxes)
//   dotnet tools/macauto.cs -- focus  <target>
//   dotnet tools/macauto.cs -- place  <target> <x> <y> <w> <h>    (in points, as macOS lays windows out)
//
// <target> is a process name (LiveGeometry.Desktop), pid:1234, window:<number> or title:substring.
// All x/y but place's are pixels relative to the top-left of the window, title bar included:
// exactly the pixel coordinates of a `shot` image (twice the points on a Retina screen).
// `shot` is the window's own picture; `--screen` is what the screen shows there, which
// includes a context menu or popup over the window (they are windows of their own).
// Needs the Accessibility and Screen Recording permissions for whatever runs it (the terminal).

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

if (args.Length == 0)
{
    Console.WriteLine("usage: list | shot | click | drag | move | wheel | key | keys | text | focus | place  (see header comment)");
    return 1;
}

try
{
    switch (args[0].ToLowerInvariant())
    {
        case "list":
            List();
            break;
        case "shot":
            Shot(Resolve(args[1]), args[2], args.Contains("--screen"));
            break;
        case "focus":
            Focus(Resolve(args[1]));
            break;
        case "click":
            Click(Resolve(args[1]), int.Parse(args[2]), int.Parse(args[3]), args.Length > 4 && !args[4].StartsWith("--") ? args[4] : "left", ModifierFlags(args));
            break;
        case "move":
        {
            var window = Resolve(args[1]);
            Focus(window);
            Post(MouseEvent(MouseMoved, ToScreen(window, int.Parse(args[2]), int.Parse(args[3])), 0));
            break;
        }
        case "drag":
        {
            int steps = args.Length > 6 && !args[6].StartsWith("--") ? int.Parse(args[6]) : 12;
            Drag(Resolve(args[1]), int.Parse(args[2]), int.Parse(args[3]), int.Parse(args[4]), int.Parse(args[5]), steps, ModifierFlags(args));
            break;
        }
        case "wheel":
            Wheel(Resolve(args[1]), int.Parse(args[2]), int.Parse(args[3]), int.Parse(args[4]));
            break;
        case "key":
        {
            Focus(Resolve(args[1]));
            var modifiers = args.Skip(3).Select(m => m.ToLowerInvariant()).ToArray();
            ulong flags = (modifiers.Contains("cmd") ? Command : 0) | (modifiers.Contains("ctrl") ? Control : 0)
                | (modifiers.Contains("shift") ? Shift : 0) | (modifiers.Contains("alt") ? Alt : 0);
            PressKey(KeyCode(args[2]), flags);
            break;
        }
        case "keys":
            Focus(Resolve(args[1]));
            foreach (var c in args[2])
            {
                PressKey(KeyCode(c.ToString()), char.IsUpper(c) ? Shift : 0);
            }

            break;
        case "text":
            Focus(Resolve(args[1]));
            foreach (var c in args[2])
            {
                TypeUnicode(c);
            }

            break;
        case "place":
        {
            var window = Resolve(args[1]);
            Osascript($"tell application \"System Events\" to tell (first process whose unix id is {window.Pid})",
                $"set position of window 1 to {{{args[2]}, {args[3]}}}",
                $"set size of window 1 to {{{args[4]}, {args[5]}}}",
                "end tell");
            break;
        }
        default:
            Console.WriteLine("unknown command " + args[0]);
            return 1;
    }
}
catch (Exception ex)
{
    Console.WriteLine("error: " + ex.Message);
    return 2;
}

return 0;

// ---------------------------------------------------------------- windows

static List<WindowInfo> Windows()
{
    var result = new List<WindowInfo>();
    var list = CGWindowListCopyWindowInfo(OnScreenOnly | ExcludeDesktopElements, 0);
    if (list == IntPtr.Zero)
    {
        return result;
    }

    try
    {
        long count = CFArrayGetCount(list);
        for (long i = 0; i < count; i++)
        {
            var dictionary = CFArrayGetValueAtIndex(list, i);
            var window = new WindowInfo()
            {
                Number = GetInt(dictionary, "kCGWindowNumber"),
                Pid = GetInt(dictionary, "kCGWindowOwnerPID"),
                Layer = GetInt(dictionary, "kCGWindowLayer"),
                Owner = GetString(dictionary, "kCGWindowOwnerName"),
                Title = GetString(dictionary, "kCGWindowName"),
            };
            var bounds = GetValue(dictionary, "kCGWindowBounds");
            if (bounds != IntPtr.Zero && CGRectMakeWithDictionaryRepresentation(bounds, out var rect))
            {
                window.Bounds = rect;
            }

            if (window.Bounds.Width > 1 && window.Bounds.Height > 1)
            {
                result.Add(window);
            }
        }
    }
    finally
    {
        CFRelease(list);
    }

    return result;
}

static WindowInfo Resolve(string target)
{
    var candidates = new List<WindowInfo>();
    foreach (var window in Windows())
    {
        bool match;
        if (target.StartsWith("pid:", StringComparison.OrdinalIgnoreCase))
        {
            match = window.Pid == int.Parse(target.Substring(4));
        }
        else if (target.StartsWith("window:", StringComparison.OrdinalIgnoreCase))
        {
            match = window.Number == int.Parse(target.Substring(7));
        }
        else if (target.StartsWith("title:", StringComparison.OrdinalIgnoreCase))
        {
            match = (window.Title ?? "").Contains(target.Substring(6), StringComparison.OrdinalIgnoreCase);
        }
        else
        {
            match = string.Equals(ProcessName(window.Pid), target, StringComparison.OrdinalIgnoreCase)
                || string.Equals(window.Owner, target, StringComparison.OrdinalIgnoreCase);
        }

        if (match)
        {
            candidates.Add(window);
        }
    }

    if (candidates.Count == 0)
    {
        throw new Exception($"no on-screen window matches '{target}' (try: list)");
    }

    // the list is front to back; a normal window (layer 0) is the app's own, a popup or a
    // menu floats above it on a higher layer
    return candidates.FirstOrDefault(w => w.Layer == 0) ?? candidates[0];
}

static void List()
{
    foreach (var window in Windows())
    {
        var b = window.Bounds;
        Console.WriteLine($"window:{window.Number,-6} pid:{window.Pid,-6} layer {window.Layer,-4} {b.X},{b.Y} {b.Width}x{b.Height}  {window.Owner} / {ProcessName(window.Pid)}  \"{window.Title}\"");
    }
}

static string ProcessName(int pid)
{
    try
    {
        return Process.GetProcessById(pid).ProcessName;
    }
    catch
    {
        return "?";
    }
}

static void Focus(WindowInfo window)
{
    Osascript($"tell application \"System Events\" to set frontmost of (first process whose unix id is {window.Pid}) to true");
    Thread.Sleep(150);
}

static void Osascript(params string[] lines)
{
    var info = new ProcessStartInfo("osascript") { RedirectStandardError = true };
    foreach (var line in lines)
    {
        info.ArgumentList.Add("-e");
        info.ArgumentList.Add(line);
    }

    using var process = Process.Start(info);
    var error = process.StandardError.ReadToEnd();
    process.WaitForExit();
    if (process.ExitCode != 0)
    {
        throw new Exception("osascript: " + error.Trim());
    }
}

// ---------------------------------------------------------------- pictures

static void Shot(WindowInfo window, string path, bool screen)
{
    var b = window.Bounds;
    var info = new ProcessStartInfo("screencapture") { RedirectStandardError = true };
    info.ArgumentList.Add("-x");
    if (screen)
    {
        info.ArgumentList.Add($"-R{b.X},{b.Y},{b.Width},{b.Height}");
    }
    else
    {
        info.ArgumentList.Add("-o");
        info.ArgumentList.Add("-l" + window.Number);
    }

    info.ArgumentList.Add(Path.GetFullPath(path));
    using var process = Process.Start(info);
    var error = process.StandardError.ReadToEnd();
    process.WaitForExit();
    if (process.ExitCode != 0 || !File.Exists(path))
    {
        throw new Exception("screencapture failed: " + error.Trim());
    }

    Console.WriteLine($"{path} ({b.Width * Scale(window)}x{b.Height * Scale(window)} px, scale {Scale(window)})");
}

/// <summary>Pixels per point of the display the window is on.</summary>
static double Scale(WindowInfo window)
{
    uint display;
    uint count;
    if (CGGetDisplaysWithPoint(new CGPoint(window.Bounds.X + 1, window.Bounds.Y + 1), 1, out display, out count) != 0 || count == 0)
    {
        display = CGMainDisplayID();
    }

    var mode = CGDisplayCopyDisplayMode(display);
    try
    {
        double points = CGDisplayModeGetWidth(mode);
        return points > 0 ? CGDisplayModeGetPixelWidth(mode) / points : 1;
    }
    finally
    {
        CGDisplayModeRelease(mode);
    }
}

static CGPoint ToScreen(WindowInfo window, int x, int y)
{
    double scale = Scale(window);
    return new CGPoint(window.Bounds.X + x / scale, window.Bounds.Y + y / scale);
}

// ---------------------------------------------------------------- mouse

static void Click(WindowInfo window, int x, int y, string kind, ulong flags)
{
    Focus(window);
    var point = ToScreen(window, x, y);
    Post(MouseEvent(MouseMoved, point, flags));
    PressModifiers(flags, down: true);
    Thread.Sleep(50);
    if (kind == "right")
    {
        Post(MouseEvent(RightMouseDown, point, flags, button: 1));
        Post(MouseEvent(RightMouseUp, point, flags, button: 1));
    }
    else
    {
        Post(MouseEvent(LeftMouseDown, point, flags));
        Post(MouseEvent(LeftMouseUp, point, flags));
        if (kind == "double")
        {
            Post(MouseEvent(LeftMouseDown, point, flags, clickCount: 2));
            Post(MouseEvent(LeftMouseUp, point, flags, clickCount: 2));
        }
    }

    PressModifiers(flags, down: false);
}

static void Drag(WindowInfo window, int x1, int y1, int x2, int y2, int steps, ulong flags)
{
    Focus(window);
    var start = ToScreen(window, x1, y1);
    var end = ToScreen(window, x2, y2);
    Post(MouseEvent(MouseMoved, start, flags));
    PressModifiers(flags, down: true);
    Thread.Sleep(50);
    Post(MouseEvent(LeftMouseDown, start, flags));
    for (int i = 1; i <= steps; i++)
    {
        var point = new CGPoint(start.X + (end.X - start.X) * i / steps, start.Y + (end.Y - start.Y) * i / steps);
        Post(MouseEvent(LeftMouseDragged, point, flags));
        Thread.Sleep(10);
    }

    Post(MouseEvent(LeftMouseUp, end, flags));
    PressModifiers(flags, down: false);
}

static void Wheel(WindowInfo window, int x, int y, int notches)
{
    Focus(window);
    Post(MouseEvent(MouseMoved, ToScreen(window, x, y), 0));
    Thread.Sleep(50);
    var scroll = CGEventCreateScrollWheelEvent2(IntPtr.Zero, ScrollUnitLine, 1, notches, 0, 0);
    Post(scroll);
}

static IntPtr MouseEvent(int type, CGPoint point, ulong flags, int button = 0, int clickCount = 1)
{
    var e = CGEventCreateMouseEvent(IntPtr.Zero, type, point, button);
    CGEventSetIntegerValueField(e, MouseEventClickState, clickCount);
    if (flags != 0)
    {
        CGEventSetFlags(e, flags);
    }

    return e;
}

static void Post(IntPtr e)
{
    CGEventPost(HidEventTap, e);
    CFRelease(e);
}

// ---------------------------------------------------------------- keyboard

static ulong ModifierFlags(string[] arguments)
{
    return (arguments.Contains("--shift") ? Shift : 0) | (arguments.Contains("--alt") ? Alt : 0)
        | (arguments.Contains("--cmd") ? Command : 0) | (arguments.Contains("--ctrl") ? Control : 0);
}

static void PressModifiers(ulong flags, bool down)
{
    foreach (var (flag, code) in new[] { (Command, (ushort)55), (Shift, (ushort)56), (Alt, (ushort)58), (Control, (ushort)59) })
    {
        if ((flags & flag) != 0)
        {
            var e = CGEventCreateKeyboardEvent(IntPtr.Zero, code, down);
            CGEventSetFlags(e, down ? flags : 0);
            Post(e);
        }
    }
}

static void PressKey(ushort code, ulong flags)
{
    PressModifiers(flags, down: true);
    var down = CGEventCreateKeyboardEvent(IntPtr.Zero, code, true);
    CGEventSetFlags(down, flags);
    Post(down);
    var up = CGEventCreateKeyboardEvent(IntPtr.Zero, code, false);
    CGEventSetFlags(up, flags);
    Post(up);
    PressModifiers(flags, down: false);
    Thread.Sleep(30);
}

static void TypeUnicode(char c)
{
    foreach (var keyDown in new[] { true, false })
    {
        var e = CGEventCreateKeyboardEvent(IntPtr.Zero, 0, keyDown);
        CGEventKeyboardSetUnicodeString(e, 1, new[] { c });
        Post(e);
    }

    Thread.Sleep(20);
}

/// <summary>The virtual key code of the US layout for a key's name or character.</summary>
static ushort KeyCode(string name)
{
    var named = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase)
    {
        ["Return"] = 36, ["Enter"] = 36, ["Tab"] = 48, ["Space"] = 49, ["Backspace"] = 51,
        ["Escape"] = 53, ["Esc"] = 53, ["Delete"] = 117, ["Home"] = 115, ["End"] = 119,
        ["PageUp"] = 116, ["PageDown"] = 121, ["Left"] = 123, ["Right"] = 124, ["Down"] = 125,
        ["Up"] = 126, ["F1"] = 122, ["F2"] = 120, ["F3"] = 99, ["F4"] = 118, ["F5"] = 96,
    };
    if (named.TryGetValue(name, out var code))
    {
        return code;
    }

    const string characters = "asdfhgzxcv\0bqweryt123465=97-80]ou[ip\0lj'k;\\,/nm.\0 `";
    if (name.Length == 1)
    {
        int index = characters.IndexOf(char.ToLowerInvariant(name[0]));
        if (index >= 0)
        {
            return (ushort)index;
        }
    }

    throw new Exception("no key code for '" + name + "'");
}

// ---------------------------------------------------------------- Core Foundation

static IntPtr GetValue(IntPtr dictionary, string key)
{
    var cfKey = CFStringCreateWithCString(IntPtr.Zero, key, Utf8);
    try
    {
        return CFDictionaryGetValue(dictionary, cfKey);
    }
    finally
    {
        CFRelease(cfKey);
    }
}

static int GetInt(IntPtr dictionary, string key)
{
    var number = GetValue(dictionary, key);
    return number != IntPtr.Zero && CFNumberGetValue(number, CFNumberIntType, out int value) ? value : 0;
}

static string GetString(IntPtr dictionary, string key)
{
    var text = GetValue(dictionary, key);
    if (text == IntPtr.Zero)
    {
        return null;
    }

    var buffer = new byte[1024];
    return CFStringGetCString(text, buffer, buffer.Length, Utf8) ? Encoding.UTF8.GetString(buffer, 0, Array.IndexOf(buffer, (byte)0)) : null;
}

class WindowInfo
{
    public int Number;
    public int Pid;
    public int Layer;
    public string Owner;
    public string Title;
    public CGRect Bounds;
}

[StructLayout(LayoutKind.Sequential)]
struct CGPoint
{
    public double X;
    public double Y;

    public CGPoint(double x, double y)
    {
        X = x;
        Y = y;
    }
}

[StructLayout(LayoutKind.Sequential)]
struct CGRect
{
    public double X;
    public double Y;
    public double Width;
    public double Height;
}

static partial class Program
{
    const string CoreGraphics = "/System/Library/Frameworks/CoreGraphics.framework/CoreGraphics";
    const string CoreFoundation = "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation";

    const uint OnScreenOnly = 1;
    const uint ExcludeDesktopElements = 16;
    const uint Utf8 = 0x08000100;
    const int CFNumberIntType = 9;
    const int HidEventTap = 0;
    const int ScrollUnitLine = 1;
    const int MouseEventClickState = 1;

    const int LeftMouseDown = 1;
    const int LeftMouseUp = 2;
    const int RightMouseDown = 3;
    const int RightMouseUp = 4;
    const int MouseMoved = 5;
    const int LeftMouseDragged = 6;

    const ulong Shift = 0x20000;
    const ulong Control = 0x40000;
    const ulong Alt = 0x80000;
    const ulong Command = 0x100000;

    [DllImport(CoreGraphics)]
    static extern IntPtr CGWindowListCopyWindowInfo(uint option, uint relativeToWindow);

    [DllImport(CoreGraphics)]
    static extern bool CGRectMakeWithDictionaryRepresentation(IntPtr dictionary, out CGRect rect);

    [DllImport(CoreGraphics)]
    static extern int CGGetDisplaysWithPoint(CGPoint point, uint maxDisplays, out uint displays, out uint count);

    [DllImport(CoreGraphics)]
    static extern uint CGMainDisplayID();

    [DllImport(CoreGraphics)]
    static extern IntPtr CGDisplayCopyDisplayMode(uint display);

    [DllImport(CoreGraphics)]
    static extern nuint CGDisplayModeGetWidth(IntPtr mode);

    [DllImport(CoreGraphics)]
    static extern nuint CGDisplayModeGetPixelWidth(IntPtr mode);

    [DllImport(CoreGraphics)]
    static extern void CGDisplayModeRelease(IntPtr mode);

    [DllImport(CoreGraphics)]
    static extern IntPtr CGEventCreateMouseEvent(IntPtr source, int type, CGPoint point, int button);

    [DllImport(CoreGraphics)]
    static extern IntPtr CGEventCreateKeyboardEvent(IntPtr source, ushort virtualKey, bool keyDown);

    // the non-variadic one: a variadic function can't be called through P/Invoke on arm64
    [DllImport(CoreGraphics)]
    static extern IntPtr CGEventCreateScrollWheelEvent2(IntPtr source, int units, uint wheelCount, int wheel1, int wheel2, int wheel3);

    [DllImport(CoreGraphics)]
    static extern void CGEventKeyboardSetUnicodeString(IntPtr e, nuint length, char[] text);

    [DllImport(CoreGraphics)]
    static extern void CGEventSetFlags(IntPtr e, ulong flags);

    [DllImport(CoreGraphics)]
    static extern void CGEventSetIntegerValueField(IntPtr e, int field, long value);

    [DllImport(CoreGraphics)]
    static extern void CGEventPost(int tap, IntPtr e);

    [DllImport(CoreFoundation)]
    static extern long CFArrayGetCount(IntPtr array);

    [DllImport(CoreFoundation)]
    static extern IntPtr CFArrayGetValueAtIndex(IntPtr array, long index);

    [DllImport(CoreFoundation)]
    static extern IntPtr CFDictionaryGetValue(IntPtr dictionary, IntPtr key);

    [DllImport(CoreFoundation)]
    static extern IntPtr CFStringCreateWithCString(IntPtr allocator, string text, uint encoding);

    [DllImport(CoreFoundation)]
    static extern bool CFStringGetCString(IntPtr text, byte[] buffer, long size, uint encoding);

    [DllImport(CoreFoundation)]
    static extern bool CFNumberGetValue(IntPtr number, int type, out int value);

    [DllImport(CoreFoundation)]
    static extern void CFRelease(IntPtr value);
}
