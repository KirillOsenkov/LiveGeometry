#:property Nullable=disable
#:property PublishAot=false

// webauto - drive the browser (WASM) build in headless Edge over the Chrome DevTools Protocol.
// Counterpart of winauto.cs. Needed because relinked .NET WASM does not run in every embedded
// browser, and because the Release publish (trimmed) can break in ways `dotnet run` never shows.
//
//   dotnet run tools/serve.cs -- <publish>/wwwroot [port=5005]  (separate tool: static file server, run in background)
//   dotnet run tools/webauto.cs -- start [url] [width height]  launch headless Edge (keeps running), open url
//   dotnet run tools/webauto.cs -- stop                        close that Edge
//   dotnet run tools/webauto.cs -- nav <url>
//   dotnet run tools/webauto.cs -- wait <text> [seconds=60]    wait until a console line contains text
//   dotnet run tools/webauto.cs -- console [--errors]          console output + uncaught errors since page load
//   dotnet run tools/webauto.cs -- shot <out.png>
//   dotnet run tools/webauto.cs -- click <x> <y> [left|right|double|middle]
//   dotnet run tools/webauto.cs -- drag <x1> <y1> <x2> <y2> [steps]
//   dotnet run tools/webauto.cs -- move <x> <y>
//   dotnet run tools/webauto.cs -- key <Key> [ctrl] [shift] [alt]   e.g. key Enter | key z ctrl | key F6 | key Delete
//   dotnet run tools/webauto.cs -- text <literal text>
//   dotnet run tools/webauto.cs -- eval <javascript>
//
// Coordinates are CSS pixels of the page == screenshot pixels (Edge is started with device
// scale factor 1). Edge stays alive between invocations on a fixed debugging port.

using System.Diagnostics;
using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

const int DebugPort = 9333;

// Buffers console output inside the page so it can be read by a later invocation; .NET's
// Console.WriteLine surfaces as console.log/console.error on the main thread.
const string ConsoleHook = """
(() => {
  if (window.__webauto) return;
  const log = window.__webauto = [];
  const push = (level, args) => { try { log.push(level + ': ' + Array.from(args).map(a => (a && a.stack) ? a.stack : (typeof a === 'object' ? JSON.stringify(a) : String(a))).join(' ')); } catch (e) { log.push(level + ': <unprintable>'); } };
  for (const level of ['log', 'info', 'warn', 'error', 'debug']) {
    const original = console[level];
    console[level] = function () { push(level, arguments); return original.apply(console, arguments); };
  }
  window.addEventListener('error', e => log.push('uncaught: ' + e.message + ' @ ' + e.filename + ':' + e.lineno));
  window.addEventListener('unhandledrejection', e => log.push('unhandledrejection: ' + ((e.reason && e.reason.stack) || e.reason)));
})();
""";

if (args.Length == 0)
{
    Console.WriteLine("usage: start | stop | nav | wait | console | shot | click | drag | move | key | text | eval  (see header comment)");
    return 1;
}

try
{
    switch (args[0].ToLowerInvariant())
    {
        case "start": await Start(args.Length > 1 ? args[1] : "about:blank", args.Length > 3 ? int.Parse(args[2]) : 1280, args.Length > 3 ? int.Parse(args[3]) : 800); break;
        case "stop": { using var cdp = await Cdp.Connect(DebugPort, browser: true); await cdp.Send("Browser.close"); break; }
        case "nav": { using var cdp = await Cdp.Connect(DebugPort); await Navigate(cdp, args[1]); break; }
        case "wait": await Wait(args[1], args.Length > 2 ? int.Parse(args[2]) : 60); break;
        case "console":
        {
            using var cdp = await Cdp.Connect(DebugPort);
            bool errorsOnly = args.Contains("--errors");
            foreach (var line in await ConsoleLines(cdp))
            {
                if (!errorsOnly || line.StartsWith("error") || line.StartsWith("uncaught") || line.StartsWith("unhandled") || line.Contains("CRASH"))
                {
                    Console.WriteLine(line);
                }
            }

            break;
        }
        case "eval":
        {
            using var cdp = await Cdp.Connect(DebugPort);
            Console.WriteLine(await Evaluate(cdp, args[1]));
            break;
        }
        case "shot":
        {
            using var cdp = await Cdp.Connect(DebugPort);
            var result = await cdp.Send("Page.captureScreenshot", new JsonObject { ["format"] = "png" });
            var path = Path.GetFullPath(args[1]);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllBytes(path, Convert.FromBase64String((string)result["data"]));
            Console.WriteLine("saved " + path + " (image pixels == click coordinates)");
            break;
        }
        case "move": { using var cdp = await Cdp.Connect(DebugPort); await Mouse(cdp, "mouseMoved", double.Parse(args[1]), double.Parse(args[2]), "none", 0, 0); break; }
        case "click":
        {
            using var cdp = await Cdp.Connect(DebugPort);
            double x = double.Parse(args[1]), y = double.Parse(args[2]);
            string kind = args.Length > 3 ? args[3].ToLowerInvariant() : "left";
            string button = kind is "right" or "middle" ? kind : "left";
            await Mouse(cdp, "mouseMoved", x, y, "none", 0, 0);
            for (int i = 1; i <= (kind == "double" ? 2 : 1); i++)
            {
                await Mouse(cdp, "mousePressed", x, y, button, ButtonMask(button), i);
                await Mouse(cdp, "mouseReleased", x, y, button, 0, i);
            }

            break;
        }
        case "drag":
        {
            using var cdp = await Cdp.Connect(DebugPort);
            double x1 = double.Parse(args[1]), y1 = double.Parse(args[2]), x2 = double.Parse(args[3]), y2 = double.Parse(args[4]);
            int steps = args.Length > 5 ? int.Parse(args[5]) : 12;
            await Mouse(cdp, "mouseMoved", x1, y1, "none", 0, 0);
            await Mouse(cdp, "mousePressed", x1, y1, "left", 1, 1);
            for (int i = 1; i <= steps; i++)
            {
                await Mouse(cdp, "mouseMoved", x1 + (x2 - x1) * i / steps, y1 + (y2 - y1) * i / steps, "left", 1, 0);
                await Task.Delay(15);
            }

            await Mouse(cdp, "mouseReleased", x2, y2, "left", 0, 1);
            break;
        }
        case "key": { using var cdp = await Cdp.Connect(DebugPort); await Key(cdp, args[1], args.Skip(2).ToArray()); break; }
        case "text": { using var cdp = await Cdp.Connect(DebugPort); await cdp.Send("Input.insertText", new JsonObject { ["text"] = args[1] }); break; }
        default: Console.WriteLine("unknown command " + args[0]); return 1;
    }
}
catch (Exception ex)
{
    Console.WriteLine("error: " + ex.Message);
    return 2;
}

return 0;

// ---------------------------------------------------------------- browser lifetime

static async Task Start(string url, int width, int height)
{
    if (await Cdp.IsAlive(DebugPort))
    {
        Console.WriteLine($"Edge already running on port {DebugPort}; navigating.");
    }
    else
    {
        var edge = new[]
        {
            Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
            Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
        }.FirstOrDefault(File.Exists) ?? throw new Exception("msedge.exe not found");

        var profile = Path.Combine(Path.GetTempPath(), "webauto-edge-profile");
        var info = new ProcessStartInfo(edge)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true, // keep the browser from holding our console handles open
            RedirectStandardError = true,
        };
        foreach (var a in new[]
        {
            "--headless=new", $"--remote-debugging-port={DebugPort}", $"--user-data-dir={profile}",
            $"--window-size={width},{height}", "--force-device-scale-factor=1", "--hide-scrollbars",
            "--no-first-run", "--no-default-browser-check", "--disable-extensions", "--disable-sync",
            // guest: otherwise Edge signs the throwaway profile into the Windows account and starts syncing
            "--guest", "about:blank",
        })
        {
            info.ArgumentList.Add(a);
        }

        Process.Start(info);
        for (int i = 0; i < 50 && !await Cdp.IsAlive(DebugPort); i++)
        {
            await Task.Delay(200);
        }
    }

    using var cdp = await Cdp.Connect(DebugPort);
    await Navigate(cdp, url);
    Console.WriteLine($"ready: {url}");
}

static async Task Navigate(Cdp cdp, string url)
{
    await cdp.Send("Page.enable");
    // Registered scripts die with this CDP session, so the hook must be in place for the
    // navigation we trigger right now; later reloads by the page itself are re-hooked lazily.
    await cdp.Send("Page.addScriptToEvaluateOnNewDocument", new JsonObject { ["source"] = ConsoleHook });
    await cdp.Send("Page.navigate", new JsonObject { ["url"] = url });
    await cdp.WaitForEvent("Page.loadEventFired", TimeSpan.FromSeconds(30));
}

static async Task Wait(string text, int seconds)
{
    using var cdp = await Cdp.Connect(DebugPort);
    var deadline = DateTime.UtcNow.AddSeconds(seconds);
    while (DateTime.UtcNow < deadline)
    {
        var hit = (await ConsoleLines(cdp)).FirstOrDefault(l => l.Contains(text, StringComparison.OrdinalIgnoreCase));
        if (hit != null)
        {
            Console.WriteLine("found: " + hit);
            return;
        }

        await Task.Delay(500);
    }

    throw new Exception($"'{text}' did not appear in the console within {seconds}s");
}

static async Task<string[]> ConsoleLines(Cdp cdp)
{
    var json = await Evaluate(cdp, "JSON.stringify(window.__webauto || ['(console hook not installed: page was not opened via start/nav)'])");
    return JsonSerializer.Deserialize<string[]>(json);
}

static async Task<string> Evaluate(Cdp cdp, string expression)
{
    var result = await cdp.Send("Runtime.evaluate", new JsonObject { ["expression"] = expression, ["returnByValue"] = true, ["awaitPromise"] = true });
    if (result["exceptionDetails"] != null)
    {
        throw new Exception("js: " + (result["exceptionDetails"]["exception"]?["description"] ?? result["exceptionDetails"]["text"]));
    }

    var value = result["result"]?["value"];
    return value == null ? "undefined" : value.GetValueKind() == JsonValueKind.String ? (string)value : value.ToJsonString();
}

// ---------------------------------------------------------------- input

static int ButtonMask(string button) => button switch { "left" => 1, "right" => 2, "middle" => 4, _ => 0 };

static Task<JsonNode> Mouse(Cdp cdp, string type, double x, double y, string button, int buttons, int clickCount) =>
    cdp.Send("Input.dispatchMouseEvent", new JsonObject
    {
        ["type"] = type, ["x"] = x, ["y"] = y, ["button"] = button, ["buttons"] = buttons, ["clickCount"] = clickCount,
    });

static async Task Key(Cdp cdp, string key, string[] modifierNames)
{
    int modifiers = 0;
    foreach (var m in modifierNames)
    {
        modifiers |= m.ToLowerInvariant() switch { "alt" => 1, "ctrl" => 2, "meta" => 4, "shift" => 8, _ => throw new Exception("unknown modifier " + m) };
    }

    var named = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        ["Enter"] = 13, ["Tab"] = 9, ["Escape"] = 27, ["Backspace"] = 8, ["Delete"] = 46, ["Insert"] = 45, ["Home"] = 36, ["End"] = 35,
        ["PageUp"] = 33, ["PageDown"] = 34, ["ArrowUp"] = 38, ["ArrowDown"] = 40, ["ArrowLeft"] = 37, ["ArrowRight"] = 39, [" "] = 32,
    };
    for (int f = 1; f <= 12; f++)
    {
        named["F" + f] = 111 + f;
    }

    string code;
    int vk;
    string text = null;
    if (named.TryGetValue(key, out vk))
    {
        key = named.Keys.First(k => string.Equals(k, key, StringComparison.OrdinalIgnoreCase));
        code = key == " " ? "Space" : key;
        text = key == "Enter" ? "\r" : key == " " ? " " : null;
    }
    else if (key.Length == 1)
    {
        vk = char.ToUpperInvariant(key[0]);
        code = char.IsLetter(key[0]) ? "Key" + char.ToUpperInvariant(key[0]) : char.IsDigit(key[0]) ? "Digit" + key : "";
        // with Ctrl/Alt held the key is a shortcut, not text input
        text = (modifiers & 0b0111) == 0 ? key : null;
    }
    else
    {
        throw new Exception("unknown key " + key + " (use DOM key names: Enter, Escape, ArrowLeft, F6, Delete, or a single character)");
    }

    JsonObject Event(string type) => new JsonObject
    {
        ["type"] = type, ["key"] = key, ["code"] = code, ["windowsVirtualKeyCode"] = vk, ["nativeVirtualKeyCode"] = vk, ["modifiers"] = modifiers,
    };

    var down = Event(text != null ? "keyDown" : "rawKeyDown");
    if (text != null)
    {
        down["text"] = text;
    }

    await cdp.Send("Input.dispatchKeyEvent", down);
    await cdp.Send("Input.dispatchKeyEvent", Event("keyUp"));
}

// ---------------------------------------------------------------- CDP client

sealed class Cdp : IDisposable
{
    private readonly ClientWebSocket socket = new ClientWebSocket();
    private readonly List<JsonNode> events = new List<JsonNode>();
    private int nextId;

    public static async Task<bool> IsAlive(int port)
    {
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(1) };
            await http.GetStringAsync($"http://127.0.0.1:{port}/json/version");
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static async Task<Cdp> Connect(int port, bool browser = false)
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(3) };
        string url;
        try
        {
            if (browser)
            {
                url = (string)JsonNode.Parse(await http.GetStringAsync($"http://127.0.0.1:{port}/json/version"))["webSocketDebuggerUrl"];
            }
            else
            {
                var targets = JsonNode.Parse(await http.GetStringAsync($"http://127.0.0.1:{port}/json")).AsArray();
                // Edge likes to open its own pages (sync confirmation, what's new) as extra targets.
                var page = targets.FirstOrDefault(t => (string)t["type"] == "page" && !((string)t["url"]).StartsWith("edge://")) ?? throw new Exception("no page target");
                url = (string)page["webSocketDebuggerUrl"];
            }
        }
        catch (HttpRequestException)
        {
            throw new Exception($"no browser on port {port}; run: start <url>");
        }

        var cdp = new Cdp();
        cdp.socket.Options.KeepAliveInterval = TimeSpan.Zero;
        await cdp.socket.ConnectAsync(new Uri(url), CancellationToken.None);
        return cdp;
    }

    public async Task<JsonNode> Send(string method, JsonObject parameters = null)
    {
        int id = ++nextId;
        var message = new JsonObject { ["id"] = id, ["method"] = method, ["params"] = parameters ?? new JsonObject() };
        await socket.SendAsync(Encoding.UTF8.GetBytes(message.ToJsonString()), WebSocketMessageType.Text, true, CancellationToken.None);
        while (true)
        {
            var reply = await Receive(TimeSpan.FromSeconds(60));
            if (reply["id"] != null && (int)reply["id"] == id)
            {
                if (reply["error"] != null)
                {
                    throw new Exception($"{method}: {reply["error"]["message"]}");
                }

                return reply["result"];
            }

            events.Add(reply);
        }
    }

    public async Task WaitForEvent(string method, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (true)
        {
            if (events.Any(e => (string)e["method"] == method))
            {
                return;
            }

            var remaining = deadline - DateTime.UtcNow;
            if (remaining <= TimeSpan.Zero)
            {
                throw new Exception("timed out waiting for " + method);
            }

            events.Add(await Receive(remaining));
        }
    }

    private async Task<JsonNode> Receive(TimeSpan timeout)
    {
        using var cancel = new CancellationTokenSource(timeout);
        using var buffer = new MemoryStream();
        var chunk = new byte[1 << 16];
        while (true)
        {
            var result = await socket.ReceiveAsync(chunk, cancel.Token);
            buffer.Write(chunk, 0, result.Count);
            if (result.EndOfMessage)
            {
                return JsonNode.Parse(Encoding.UTF8.GetString(buffer.GetBuffer(), 0, (int)buffer.Length));
            }
        }
    }

    public void Dispose() => socket.Dispose();
}
