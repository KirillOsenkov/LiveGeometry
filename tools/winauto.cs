#:property TargetFramework=net10.0-windows
#:property UseWindowsForms=true
#:property AllowUnsafeBlocks=true
#:property Nullable=disable
#:property PublishAot=false

// winauto - minimal Win32 UI automation for inspecting desktop apps (the VB6 Geometry.exe,
// the Avalonia desktop app, anything with a top-level window).
//
//   dotnet tools/winauto.cs -- list
//   dotnet tools/winauto.cs -- tree   <target>
//   dotnet tools/winauto.cs -- menu   <target>
//   dotnet tools/winauto.cs -- invoke <target> <menuId>              (WM_COMMAND, no focus needed)
//   dotnet tools/winauto.cs -- shot   <target> <out.png> [--screen]
//   dotnet tools/winauto.cs -- click  <target> <x> <y> [left|right|double|middle]
//   dotnet tools/winauto.cs -- drag   <target> <x1> <y1> <x2> <y2> [steps]
//   dotnet tools/winauto.cs -- move   <target> <x> <y>
//   dotnet tools/winauto.cs -- keys   <target> <sendkeys-syntax>     e.g. "^s" "{ENTER}" "%f" "{DOWN 3}"
//   dotnet tools/winauto.cs -- text   <target> <literal text>
//   dotnet tools/winauto.cs -- focus  <target>
//   dotnet tools/winauto.cs -- place  <target> <x> <y> <w> <h>       (move/resize the window)
//
// <target> is a process name (Geometry), pid:1234, hwnd:0x1A2B, or title:substring.
// All x/y are physical pixels relative to the top-left of the window rectangle, i.e. exactly
// the pixel coordinates of a `shot` image taken without scaling.

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

SetProcessDpiAwarenessContext(new IntPtr(-4)); // per-monitor v2: physical pixels everywhere

if (args.Length == 0)
{
    Console.WriteLine("usage: list | tree | menu | invoke | shot | click | drag | move | keys | text | focus | place  (see header comment)");
    return 1;
}

try
{
    switch (args[0].ToLowerInvariant())
    {
        case "list": List(); break;
        case "tree": Tree(Resolve(args[1])); break;
        case "menu": Menu(Resolve(args[1])); break;
        case "invoke": PostMessage(Resolve(args[1]), 0x0111 /*WM_COMMAND*/, new IntPtr(int.Parse(args[2])), IntPtr.Zero); break;
        case "shot": Shot(Resolve(args[1]), args[2], args.Contains("--screen")); break;
        case "focus": Focus(Resolve(args[1])); break;
        case "click": Click(Resolve(args[1]), int.Parse(args[2]), int.Parse(args[3]), args.Length > 4 ? args[4] : "left"); break;
        case "move": { var h = Resolve(args[1]); Focus(h); MoveTo(h, int.Parse(args[2]), int.Parse(args[3])); break; }
        case "drag": Drag(Resolve(args[1]), int.Parse(args[2]), int.Parse(args[3]), int.Parse(args[4]), int.Parse(args[5]), args.Length > 6 ? int.Parse(args[6]) : 12); break;
        case "keys": { Focus(Resolve(args[1])); SendKeysSyntax(args[2]); break; }
        case "text": { Focus(Resolve(args[1])); foreach (var c in args[2]) Unicode(c); break; }
        case "place": { var h = Resolve(args[1]); ShowWindow(h, 9); MoveWindow(h, int.Parse(args[2]), int.Parse(args[3]), int.Parse(args[4]), int.Parse(args[5]), true); break; }
        default: Console.WriteLine("unknown command " + args[0]); return 1;
    }
}
catch (Exception ex)
{
    Console.WriteLine("error: " + ex.Message);
    return 2;
}

return 0;

// ---------------------------------------------------------------- targets

static IntPtr Resolve(string target)
{
    if (target.StartsWith("hwnd:", StringComparison.OrdinalIgnoreCase))
    {
        var s = target.Substring(5);
        return new IntPtr(s.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? Convert.ToInt64(s.Substring(2), 16) : long.Parse(s));
    }

    var candidates = new List<IntPtr>();
    foreach (var w in TopLevel())
    {
        GetWindowThreadProcessId(w, out var pid);
        bool match;
        if (target.StartsWith("pid:", StringComparison.OrdinalIgnoreCase))
        {
            match = pid == uint.Parse(target.Substring(4));
        }
        else if (target.StartsWith("title:", StringComparison.OrdinalIgnoreCase))
        {
            match = Text(w).Contains(target.Substring(6), StringComparison.OrdinalIgnoreCase);
        }
        else
        {
            match = string.Equals(ProcessName(pid), target, StringComparison.OrdinalIgnoreCase);
        }

        if (match)
        {
            candidates.Add(w);
        }
    }

    if (candidates.Count == 0)
    {
        throw new Exception($"no visible top-level window matches '{target}' (try: list)");
    }

    // Prefer the foreground-most window that has no owner... but a modal dialog is what the
    // user sees, so prefer an enabled window: a form with a modal child is disabled.
    foreach (var w in candidates)
    {
        if (IsWindowEnabled(w))
        {
            return w; // EnumWindows is z-ordered, so this is the topmost enabled match
        }
    }

    return candidates[0];
}

static List<IntPtr> TopLevel()
{
    var result = new List<IntPtr>();
    EnumWindows((h, _) =>
    {
        if (IsWindowVisible(h) && GetWindowRect(h, out var r) && r.Right > r.Left && r.Bottom > r.Top)
        {
            result.Add(h);
        }

        return true;
    }, IntPtr.Zero);
    return result;
}

static string ProcessName(uint pid)
{
    try { return Process.GetProcessById((int)pid).ProcessName; } catch { return "?"; }
}

static string Text(IntPtr h)
{
    // WM_GETTEXT with a timeout: works for child controls of other processes and never hangs.
    var buffer = Marshal.AllocHGlobal(2048);
    try
    {
        Marshal.WriteInt16(buffer, 0);
        SendMessageTimeout(h, 0x000D, new IntPtr(1024), buffer, 2, 200, out _);
        return Marshal.PtrToStringUni(buffer) ?? "";
    }
    finally
    {
        Marshal.FreeHGlobal(buffer);
    }
}

static string Class(IntPtr h)
{
    var sb = new StringBuilder(256);
    GetClassName(h, sb, sb.Capacity);
    return sb.ToString();
}

// ---------------------------------------------------------------- inspection

static void List()
{
    foreach (var w in TopLevel())
    {
        var title = Text(w);
        if (title.Length == 0)
        {
            continue;
        }

        GetWindowThreadProcessId(w, out var pid);
        GetWindowRect(w, out var r);
        Console.WriteLine($"hwnd:0x{w.ToInt64():X} pid:{pid} {ProcessName(pid)} [{Class(w)}] \"{title}\" [{r.Left},{r.Top} {r.Right - r.Left}x{r.Bottom - r.Top}]{(IsWindowEnabled(w) ? "" : " disabled")}{(w == GetForegroundWindow() ? " FOREGROUND" : "")}");
    }
}

static void Tree(IntPtr root)
{
    GetWindowRect(root, out var origin);
    void Print(IntPtr h, int depth)
    {
        GetWindowRect(h, out var r);
        var text = Text(h).Replace("\r", "\\r").Replace("\n", "\\n");
        if (text.Length > 80)
        {
            text = text.Substring(0, 80) + "...";
        }

        // rect is relative to the root window = screenshot coordinates
        Console.WriteLine($"{new string(' ', depth * 2)}hwnd:0x{h.ToInt64():X} [{Class(h)}] \"{text}\" [{r.Left - origin.Left},{r.Top - origin.Top} {r.Right - r.Left}x{r.Bottom - r.Top}] id={GetDlgCtrlID(h)}{(IsWindowVisible(h) ? "" : " hidden")}{(IsWindowEnabled(h) ? "" : " disabled")}");
        var child = GetWindow(h, 5 /*GW_CHILD*/);
        while (child != IntPtr.Zero)
        {
            Print(child, depth + 1);
            child = GetWindow(child, 2 /*GW_HWNDNEXT*/);
        }
    }

    Print(root, 0);
}

static void Menu(IntPtr window)
{
    var menu = GetMenu(window);
    if (menu == IntPtr.Zero)
    {
        Console.WriteLine("(window has no native menu bar)");
        return;
    }

    void Print(IntPtr m, int depth)
    {
        int count = GetMenuItemCount(m);
        for (int i = 0; i < count; i++)
        {
            var sb = new StringBuilder(512);
            GetMenuString(m, (uint)i, sb, sb.Capacity, 0x400 /*MF_BYPOSITION*/);
            uint state = GetMenuState(m, (uint)i, 0x400);
            var sub = GetSubMenu(m, i);
            uint id = sub == IntPtr.Zero ? GetMenuItemID(m, i) : 0;
            string label = sb.Length == 0 && (state & 0x800) != 0 ? "--------" : sb.ToString().Replace("\t", "  \\t ");
            string flags = ((state & 0x3) != 0 ? " disabled" : "") + ((state & 0x8) != 0 ? " checked" : "");
            Console.WriteLine($"{new string(' ', depth * 2)}{label}{(sub == IntPtr.Zero && sb.Length > 0 ? $"  (id={id})" : "")}{flags}");
            if (sub != IntPtr.Zero)
            {
                Print(sub, depth + 1);
            }
        }
    }

    Print(menu, 0);
}

static void Shot(IntPtr h, string path, bool fromScreen)
{
    GetWindowRect(h, out var r);
    int w = r.Right - r.Left, ht = r.Bottom - r.Top;
    using var bmp = new Bitmap(w, ht, PixelFormat.Format32bppArgb);
    using (var g = Graphics.FromImage(bmp))
    {
        bool ok = false;
        if (!fromScreen)
        {
            // A DPI-unaware app (VB6) on a scaled monitor paints at 96 DPI and DWM stretches
            // it; PrintWindow yields that small unscaled image. Render it at its logical
            // size and stretch it ourselves so image pixels stay equal to click coordinates.
            uint windowDpi = GetDpiForWindow(h);
            GetDpiForMonitor(MonitorFromWindow(h, 2), 0, out uint monitorDpi, out _);
            int lw = w, lh = ht;
            if (windowDpi != 0 && monitorDpi != 0 && windowDpi != monitorDpi)
            {
                lw = (int)Math.Ceiling(w * (double)windowDpi / monitorDpi);
                lh = (int)Math.Ceiling(ht * (double)windowDpi / monitorDpi);
            }

            using var logical = new Bitmap(lw, lh, PixelFormat.Format32bppArgb);
            using (var lg = Graphics.FromImage(logical))
            {
                var dc = lg.GetHdc();
                ok = PrintWindow(h, dc, 2 /*PW_RENDERFULLCONTENT*/);
                lg.ReleaseHdc(dc);
            }

            if (ok)
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
                g.DrawImage(logical, new Rectangle(0, 0, w, ht));
            }
        }

        if (!ok)
        {
            Focus(h);
            Thread.Sleep(250);
            g.CopyFromScreen(r.Left, r.Top, 0, 0, new Size(w, ht));
        }
    }

    Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
    bmp.Save(path, ImageFormat.Png);
    Console.WriteLine($"saved {path} {w}x{ht} (window at {r.Left},{r.Top}; image pixels == click coordinates)");
}

// ---------------------------------------------------------------- input

static void Focus(IntPtr h)
{
    if (IsIconic(h))
    {
        ShowWindow(h, 9 /*SW_RESTORE*/);
    }

    if (GetForegroundWindow() == h)
    {
        return;
    }

    // A background process may not steal the foreground unless it "owns" the last input
    // event; a synthetic ALT tap satisfies that rule.
    Key(0x12, false); Key(0x12, true);
    SetForegroundWindow(h);
    BringWindowToTop(h);
    Thread.Sleep(150);
    if (GetForegroundWindow() != h)
    {
        Console.WriteLine("warning: could not bring the window to the foreground; input may go astray");
    }
}

static void MoveTo(IntPtr h, int x, int y)
{
    GetWindowRect(h, out var r);
    int sx = r.Left + x, sy = r.Top + y;
    // absolute coordinates are normalized to 0..65535 over the virtual desktop
    int vx = GetSystemMetrics(76), vy = GetSystemMetrics(77), vw = GetSystemMetrics(78), vh = GetSystemMetrics(79);
    var input = new INPUT { type = 0 };
    input.u.mi.dx = (int)Math.Round((sx - vx) * 65535.0 / (vw - 1));
    input.u.mi.dy = (int)Math.Round((sy - vy) * 65535.0 / (vh - 1));
    input.u.mi.dwFlags = 0x0001 /*MOVE*/ | 0x8000 /*ABSOLUTE*/ | 0x4000 /*VIRTUALDESK*/;
    SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
    Thread.Sleep(20);
}

static void Button(uint flags)
{
    var input = new INPUT { type = 0 };
    input.u.mi.dwFlags = flags;
    SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
    Thread.Sleep(30);
}

static void Click(IntPtr h, int x, int y, string kind)
{
    Focus(h);
    MoveTo(h, x, y);
    switch (kind.ToLowerInvariant())
    {
        case "right": Button(0x0008); Button(0x0010); break;
        case "middle": Button(0x0020); Button(0x0040); break;
        case "double": Button(0x0002); Button(0x0004); Button(0x0002); Button(0x0004); break;
        default: Button(0x0002); Button(0x0004); break;
    }
}

static void Drag(IntPtr h, int x1, int y1, int x2, int y2, int steps)
{
    Focus(h);
    MoveTo(h, x1, y1);
    Button(0x0002);
    for (int i = 1; i <= steps; i++)
    {
        MoveTo(h, x1 + (x2 - x1) * i / steps, y1 + (y2 - y1) * i / steps);
    }

    Thread.Sleep(50);
    Button(0x0004);
}

static void Key(ushort vk, bool up)
{
    var input = new INPUT { type = 1 };
    input.u.ki.wVk = vk;
    // navigation keys need the extended flag or they are read as numpad keys
    bool extended = vk is >= 0x21 and <= 0x28 or 0x2D or 0x2E;
    input.u.ki.dwFlags = (up ? 0x0002u : 0) | (extended ? 0x0001u : 0);
    SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
    Thread.Sleep(15);
}

static void Unicode(char c)
{
    if (c == '\n') { Key(0x0D, false); Key(0x0D, true); return; }
    foreach (var up in new[] { false, true })
    {
        var input = new INPUT { type = 1 };
        input.u.ki.wScan = c;
        input.u.ki.dwFlags = 0x0004u /*UNICODE*/ | (up ? 0x0002u : 0);
        SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
    }

    Thread.Sleep(10);
}

static void SendKeysSyntax(string keys)
{
    var named = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase)
    {
        ["ENTER"] = 0x0D, ["TAB"] = 0x09, ["ESC"] = 0x1B, ["ESCAPE"] = 0x1B, ["BACKSPACE"] = 0x08, ["BS"] = 0x08,
        ["DELETE"] = 0x2E, ["DEL"] = 0x2E, ["INSERT"] = 0x2D, ["HOME"] = 0x24, ["END"] = 0x23, ["PGUP"] = 0x21, ["PGDN"] = 0x22,
        ["UP"] = 0x26, ["DOWN"] = 0x28, ["LEFT"] = 0x25, ["RIGHT"] = 0x27, ["SPACE"] = 0x20, ["APPS"] = 0x5D,
    };
    for (int f = 1; f <= 24; f++)
    {
        named["F" + f] = (ushort)(0x70 + f - 1);
    }

    var held = new List<ushort>();
    void Release()
    {
        for (int i = held.Count - 1; i >= 0; i--) Key(held[i], true);
        held.Clear();
    }

    void Press(char c)
    {
        if (char.IsAsciiLetterOrDigit(c) && (held.Count > 0 || !char.IsAsciiLetterUpper(c)))
        {
            // A real virtual key, not a unicode packet: shortcuts (Ctrl+S, or a plain "w"
            // handled in KeyDown/KeyUp) only see virtual keys. Assumes a Latin layout.
            ushort vk = char.ToUpperInvariant(c);
            Key(vk, false); Key(vk, true);
        }
        else
        {
            Unicode(c);
        }
    }

    for (int i = 0; i < keys.Length; i++)
    {
        char c = keys[i];
        if (c == '+') { held.Add(0x10); Key(0x10, false); continue; }
        if (c == '^') { held.Add(0x11); Key(0x11, false); continue; }
        if (c == '%') { held.Add(0x12); Key(0x12, false); continue; }
        if (c == '~') { Key(0x0D, false); Key(0x0D, true); Release(); continue; }
        if (c == '(')
        {
            int close = keys.IndexOf(')', i);
            foreach (var g in keys.Substring(i + 1, close - i - 1)) Press(g);
            i = close;
            Release();
            continue;
        }

        if (c == '{')
        {
            int close = keys.IndexOf('}', i + 2 <= keys.Length ? i + 2 : i + 1);
            var body = keys.Substring(i + 1, close - i - 1);
            i = close;
            var parts = body.Split(' ', 2);
            int repeat = parts.Length > 1 && int.TryParse(parts[1], out var n) ? n : 1;
            for (int k = 0; k < repeat; k++)
            {
                if (named.TryGetValue(parts[0], out var vk)) { Key(vk, false); Key(vk, true); }
                else foreach (var g in parts[0]) Unicode(g);
            }

            Release();
            continue;
        }

        Press(c);
        Release();
    }

    Release();
}

// ---------------------------------------------------------------- interop

[DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr value);
[DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
[DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr h);
[DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
[DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT rect);
[DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder text, int max);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessageTimeout(IntPtr h, uint msg, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);
[DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint msg, IntPtr wParam, IntPtr lParam);
[DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr h, uint cmd);
[DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr h);
[DllImport("user32.dll")] static extern IntPtr GetMenu(IntPtr h);
[DllImport("user32.dll")] static extern int GetMenuItemCount(IntPtr menu);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetMenuString(IntPtr menu, uint item, StringBuilder text, int max, uint flags);
[DllImport("user32.dll")] static extern uint GetMenuState(IntPtr menu, uint item, uint flags);
[DllImport("user32.dll")] static extern IntPtr GetSubMenu(IntPtr menu, int pos);
[DllImport("user32.dll")] static extern uint GetMenuItemID(IntPtr menu, int pos);
[DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
[DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr h);
[DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
[DllImport("user32.dll")] static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool repaint);
[DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
[DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr h);
[DllImport("user32.dll")] static extern IntPtr MonitorFromWindow(IntPtr h, uint flags);
[DllImport("shcore.dll")] static extern int GetDpiForMonitor(IntPtr monitor, int type, out uint dpiX, out uint dpiY);
[DllImport("user32.dll")] static extern uint SendInput(uint count, INPUT[] inputs, int size);

delegate bool EnumWindowsProc(IntPtr h, IntPtr lParam);

[StructLayout(LayoutKind.Sequential)]
struct RECT { public int Left, Top, Right, Bottom; }

[StructLayout(LayoutKind.Sequential)]
struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr dwExtraInfo; }

[StructLayout(LayoutKind.Sequential)]
struct KEYBDINPUT { public ushort wVk, wScan; public uint dwFlags, time; public IntPtr dwExtraInfo; }

[StructLayout(LayoutKind.Explicit)]
struct INPUTUNION
{
    [FieldOffset(0)] public MOUSEINPUT mi;
    [FieldOffset(0)] public KEYBDINPUT ki;
}

[StructLayout(LayoutKind.Sequential)]
struct INPUT { public uint type; public INPUTUNION u; }
