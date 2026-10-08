#:property Nullable=disable
#:property PublishAot=false

// loadperf - how a page loads, as F12 shows it: headless Edge over the Chrome DevTools Protocol
// (port 9334, apart from webauto's), twice - a first visit (a new profile, nothing cached) and a
// returning visitor's (a new browser on the same profile, everything cached). For each: the
// first byte, the splash, the downloads, the console lines with their times (the app's build
// line says when .NET runs), Avalonia's first frame (it adds splash-close then) and when the
// app took the splash down (app-ready: the tiles in view loaded, or a drawing laid out), and
// the long tasks of the page's thread (each load of a gallery tile is one). All of it also goes
// to first-visit.json and returning-visit.json.
//
//   dotnet run tools/loadperf.cs -- <url> <out folder> [--seconds 20] [--size 1700x1000] [--trace] [--phone]
//
// --seconds  how long each visit is watched
// --trace    the first visit recorded as F12's Performance panel records one: trace.json (open it
//            there with Load profile) and filmstrip/, what was on screen every half second
//            (`dotnet tools/contactsheet.cs -- <out folder>/filmstrip <sheet.png> 4 500` puts it
//            on one image). Recording slows the page down a little.
// --phone    a phone: 390x645 at 3x, the CPU 4 times slower, DevTools' "Fast 4G"
// --wifi     with --phone: no network throttling (a local server has no brotli, so the "Fast 4G"
//            download of a local publish is not what the site's would be; the CPU part is)
//
// Each visit gets a browser of its own, closed before the next starts: another one running
// (even one idle in the background) takes CPU from the page being measured.

using System.Diagnostics;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

const int Port = 9334;

// in the page before its own scripts run: the console lines and the long tasks with their
// times (milliseconds since the navigation began), and when the splash closed
const string Hook = """
(() => {
  const record = window.__loadperf = { console: [], longTasks: [], splashClosed: null, appReady: null };
  for (const level of ['log', 'info', 'warn', 'error']) {
    const original = console[level];
    console[level] = function () {
      try { record.console.push([Math.round(performance.now()), level, Array.from(arguments).map(String).join(' ')]); } catch (e) { }
      return original.apply(console, arguments);
    };
  }
  try {
    new PerformanceObserver(list => {
      for (const task of list.getEntries()) record.longTasks.push([Math.round(task.startTime), Math.round(task.duration)]);
    }).observe({ type: 'longtask', buffered: true });
  } catch (e) { }
  new MutationObserver(() => {
    if (record.splashClosed === null && document.querySelector('.avalonia-splash.splash-close')) {
      record.splashClosed = Math.round(performance.now());
    }
    if (record.appReady === null && document.querySelector('.avalonia-splash.app-ready')) {
      record.appReady = Math.round(performance.now());
    }
  }).observe(document, { attributes: true, attributeFilter: ['class'], subtree: true });
})();
""";

const string Summary = """
(() => {
  const navigation = performance.getEntriesByType('navigation')[0];
  const paint = Object.fromEntries(performance.getEntriesByType('paint').map(entry => [entry.name, Math.round(entry.startTime)]));
  const resources = performance.getEntriesByType('resource').map(resource => ({
    name: resource.name.replace(location.origin, ''), start: Math.round(resource.startTime), end: Math.round(resource.responseEnd),
    transfer: resource.transferSize, decoded: resource.decodedBodySize }));
  const record = window.__loadperf || {};
  return JSON.stringify({
    navigation: navigation && {
      firstByte: Math.round(navigation.responseStart), domContentLoaded: Math.round(navigation.domContentLoadedEventEnd),
      load: Math.round(navigation.loadEventEnd), transfer: navigation.transferSize },
    paint, splashClosed: record.splashClosed, appReady: record.appReady, resources, console: record.console || [], longTasks: record.longTasks || [] });
})()
""";

if (args.Length < 2)
{
    Console.WriteLine("usage: loadperf <url> <out folder> [--seconds 20] [--size 1700x1000] [--trace] [--phone]");
    return 1;
}

var url = args[0];
var outputFolder = Path.GetFullPath(args[1]);
var watched = TimeSpan.FromSeconds(int.Parse(ArgumentValue("--seconds") ?? "20"));
var size = (ArgumentValue("--size") ?? "1700x1000").Split('x').Select(int.Parse).ToArray();
bool trace = args.Contains("--trace");
bool phone = args.Contains("--phone");
bool wifi = args.Contains("--wifi");
Directory.CreateDirectory(outputFolder);
var profile = Path.Combine(Path.GetTempPath(), "loadperf-" + Guid.NewGuid().ToString("N"));

// one left on the port by a run that was cut short
await CloseBrowser();
try
{
    // a first visit: the profile is new, nothing is cached
    await StartBrowser();
    string firstVisit;
    using (var page = await OpenPage())
    {
        if (trace)
        {
            await StartTracing(page);
        }

        await page.Send("Page.navigate", new JsonObject { ["url"] = url });
        await Task.Delay(watched);
        firstVisit = await Evaluate(page, Summary);
        if (trace)
        {
            var tracePath = Path.Combine(outputFolder, "trace.json");
            await SaveTrace(page, tracePath);
            SaveFilmstrip(tracePath, Path.Combine(outputFolder, "filmstrip"));
        }
    }

    await CloseBrowser();

    // a returning visitor: a new browser on the same profile, whose cache has everything
    await StartBrowser();
    string returningVisit;
    using (var page = await OpenPage())
    {
        await page.Send("Page.navigate", new JsonObject { ["url"] = url });
        await Task.Delay(watched);
        returningVisit = await Evaluate(page, Summary);
    }

    File.WriteAllText(Path.Combine(outputFolder, "first-visit.json"), firstVisit);
    File.WriteAllText(Path.Combine(outputFolder, "returning-visit.json"), returningVisit);
    Report("first visit (nothing cached)", firstVisit);
    Report("returning visitor (a new browser, everything cached)", returningVisit);
}
finally
{
    await CloseBrowser();
    try
    {
        Directory.Delete(profile, recursive: true);
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
    {
        // a file still held: the profile stays in the temp folder
    }
}

return 0;

string ArgumentValue(string name)
{
    int index = Array.IndexOf(args, name);
    return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
}

async Task StartBrowser()
{
    var edge = new[]
    {
        Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
    }.FirstOrDefault(File.Exists) ?? throw new Exception("msedge.exe not found");
    // through the shell, so that the browser inherits none of our handles (see webauto.cs)
    var info = new ProcessStartInfo(edge) { UseShellExecute = true };
    foreach (var argument in new[]
    {
        "--headless=new", $"--remote-debugging-port={Port}", $"--user-data-dir={profile}",
        $"--window-size={size[0]},{size[1]}", "--force-device-scale-factor=1", "--hide-scrollbars",
        // not --guest, as webauto has it: a guest's cache is gone when the browser closes
        "--no-first-run", "--no-default-browser-check", "--disable-extensions", "--disable-sync", "about:blank",
    })
    {
        info.ArgumentList.Add(argument);
    }

    Process.Start(info);
    for (int i = 0; i < 50 && !await Cdp.IsAlive(Port); i++)
    {
        await Task.Delay(TimeSpan.FromMilliseconds(200));
    }
}

// the page, ready for a navigation: the phone's screen, CPU and network if asked for, and the hook
async Task<Cdp> OpenPage()
{
    var page = await Cdp.Connect(Port);
    await page.Send("Page.enable");
    if (phone)
    {
        await page.Send("Network.enable");
        await page.Send("Emulation.setDeviceMetricsOverride", new JsonObject
        {
            ["width"] = 390,
            ["height"] = 645,
            ["deviceScaleFactor"] = 3,
            ["mobile"] = true,
        });
        await page.Send("Emulation.setCPUThrottlingRate", new JsonObject { ["rate"] = 4 });

        // "Fast 4G": 9 Mbit/s down, 1.5 Mbit/s up, 60 ms
        if (!wifi)
        {
            await page.Send("Network.emulateNetworkConditions", new JsonObject
            {
                ["offline"] = false,
                ["latency"] = 60,
                ["downloadThroughput"] = 9_000_000 / 8,
                ["uploadThroughput"] = 1_500_000 / 8,
            });
        }
    }

    await page.Send("Page.addScriptToEvaluateOnNewDocument", new JsonObject { ["source"] = Hook });
    return page;
}

// Told to close, and waited for on its port: the process started is not the browser
// (msedge.exe hands over to another and exits), and a browser left running took CPU from
// every visit measured after it - one whose page drew without end, most of all.
static async Task CloseBrowser()
{
    if (!await Cdp.IsAlive(Port))
    {
        return;
    }

    try
    {
        using var connection = await Cdp.Connect(Port, browser: true);
        await connection.Send("Browser.close");
    }
    catch (Exception ex) when (ex is WebSocketException or HttpRequestException or OperationCanceledException)
    {
        // it may close before it answers
    }

    for (int i = 0; i < 50 && await Cdp.IsAlive(Port); i++)
    {
        await Task.Delay(TimeSpan.FromMilliseconds(200));
    }

    if (await Cdp.IsAlive(Port))
    {
        throw new Exception($"the browser on port {Port} did not close");
    }
}

// what F12's Performance panel records, screenshots included
static async Task StartTracing(Cdp page)
{
    var categories = new JsonArray(
        "devtools.timeline", "v8.execute", "disabled-by-default-devtools.timeline",
        "disabled-by-default-devtools.timeline.frame", "toplevel", "blink.console", "blink.user_timing",
        "latencyInfo", "disabled-by-default-devtools.timeline.stack", "disabled-by-default-v8.cpu_profiler",
        "disabled-by-default-devtools.screenshot", "loading");
    await page.Send("Tracing.start", new JsonObject
    {
        ["transferMode"] = "ReturnAsStream",
        ["traceConfig"] = new JsonObject { ["includedCategories"] = categories, ["excludedCategories"] = new JsonArray("*") },
    });
}

static async Task SaveTrace(Cdp page, string path)
{
    await page.Send("Tracing.end");
    var complete = await page.WaitForEvent("Tracing.tracingComplete", TimeSpan.FromMinutes(2));
    var handle = (string)complete["params"]["stream"];
    using (var file = File.Create(path))
    {
        while (true)
        {
            var chunk = await page.Send("IO.read", new JsonObject { ["handle"] = handle, ["size"] = 4 << 20 });
            var data = (string)chunk["data"];
            file.Write(chunk["base64Encoded"]?.GetValue<bool>() == true ? Convert.FromBase64String(data) : Encoding.UTF8.GetBytes(data));
            if (chunk["eof"]?.GetValue<bool>() == true)
            {
                break;
            }
        }
    }

    await page.Send("IO.close", new JsonObject { ["handle"] = handle });
    Console.WriteLine("trace: " + path);
}

// the trace's screenshots: what was on screen every half second after the navigation began
static void SaveFilmstrip(string tracePath, string folder)
{
    Directory.CreateDirectory(folder);
    using var stream = File.OpenRead(tracePath);
    using var document = JsonDocument.Parse(stream);
    var events = document.RootElement.TryGetProperty("traceEvents", out var list) ? list : document.RootElement;
    double? navigationStart = null;
    var screenshots = new List<(double Time, string Image)>();
    foreach (var traceEvent in events.EnumerateArray())
    {
        var name = traceEvent.TryGetProperty("name", out var nameProperty) ? nameProperty.GetString() : null;
        if (name == "navigationStart" && navigationStart == null && IsPageNavigation(traceEvent))
        {
            navigationStart = traceEvent.GetProperty("ts").GetDouble();
        }
        else if (name == "Screenshot"
            && traceEvent.TryGetProperty("args", out var arguments)
            && arguments.TryGetProperty("snapshot", out var snapshot))
        {
            screenshots.Add((traceEvent.GetProperty("ts").GetDouble(), snapshot.GetString()));
        }
    }

    if (navigationStart == null || screenshots.Count == 0)
    {
        Console.WriteLine($"filmstrip: no {(navigationStart == null ? "navigation" : "screenshots")} in the trace");
        return;
    }

    // the trace counts microseconds
    screenshots.Sort((first, second) => first.Time.CompareTo(second.Time));
    double end = (screenshots[^1].Time - navigationStart.Value) / 1000;
    for (double milliseconds = 500; milliseconds <= end + 500; milliseconds += 500)
    {
        var shown = screenshots.LastOrDefault(screenshot => (screenshot.Time - navigationStart.Value) / 1000 <= milliseconds);
        if (shown.Image != null)
        {
            File.WriteAllBytes(Path.Combine(folder, $"{milliseconds / 1000:00.0}s.jpg"), Convert.FromBase64String(shown.Image));
        }
    }

    Console.WriteLine($"filmstrip: {folder} ({screenshots.Count} screenshots over {end / 1000:0.0} s)");
}

// the main frame's navigation to the page, not the one to about:blank
static bool IsPageNavigation(JsonElement traceEvent)
{
    return traceEvent.TryGetProperty("args", out var arguments)
        && arguments.TryGetProperty("data", out var data)
        && data.TryGetProperty("isLoadingMainFrame", out var isMainFrame) && isMainFrame.GetBoolean()
        && data.TryGetProperty("documentLoaderURL", out var loader) && loader.GetString()?.StartsWith("http") == true;
}

static async Task<string> Evaluate(Cdp page, string expression)
{
    var result = await page.Send("Runtime.evaluate", new JsonObject { ["expression"] = expression, ["returnByValue"] = true });
    if (result["exceptionDetails"] != null)
    {
        throw new Exception("js: " + result["exceptionDetails"].ToJsonString());
    }

    return (string)result["result"]["value"];
}

static void Report(string title, string json)
{
    var root = JsonNode.Parse(json);
    var navigation = root["navigation"];
    Console.WriteLine($"== {title} ==");
    Console.WriteLine($"first byte {navigation?["firstByte"]} ms, DOMContentLoaded {navigation?["domContentLoaded"]} ms, load {navigation?["load"]} ms");
    Console.WriteLine($"first paint (the splash) {root["paint"]?["first-contentful-paint"]} ms, Avalonia's first frame {root["splashClosed"]} ms, the splash taken down (the app says it is ready) {root["appReady"]} ms");
    var resources = root["resources"].AsArray();
    long transferred = resources.Sum(resource => (long)resource["transfer"]) + (long)(navigation?["transfer"] ?? 0);
    long decoded = resources.Sum(resource => (long)resource["decoded"]);
    int lastDone = resources.Count > 0 ? resources.Max(resource => (int)resource["end"]) : 0;
    Console.WriteLine($"requests {resources.Count + 1}, transferred {transferred / 1024} KB ({decoded / 1024} KB unpacked), the last done at {lastDone} ms");
    foreach (var resource in resources.OrderByDescending(resource => (long)resource["transfer"]).Take(8))
    {
        Console.WriteLine($"  {(long)resource["transfer"] / 1024,6} KB  {resource["start"],6}-{resource["end"],-6} ms  {resource["name"]}");
    }

    var font = resources.FirstOrDefault(resource => ((string)resource["name"]).Contains("Twemoji"));
    if (font != null)
    {
        Console.WriteLine($"emoji font asked for at {font["start"]} ms, here at {font["end"]} ms ({(long)font["transfer"] / 1024} KB)");
    }

    foreach (var line in root["console"].AsArray())
    {
        Console.WriteLine($"  console {line[0],6} ms  {Shorten((string)line[2], length: 100)}");
    }

    var tasks = root["longTasks"].AsArray().Select(task => (Start: (int)task[0], Duration: (int)task[1])).ToList();
    int lastEnd = tasks.Count > 0 ? tasks.Max(task => task.Start + task.Duration) : 0;
    Console.WriteLine($"long tasks: {tasks.Count}, {tasks.Sum(task => task.Duration)} ms in all, the last ending at {lastEnd} ms");
    Console.WriteLine("  " + string.Join(" ", tasks.Select(task => $"{task.Start}+{task.Duration}")));
    Console.WriteLine();
}

static string Shorten(string text, int length)
{
    return text.Length <= length ? text : text.Substring(0, length) + "...";
}

// ---------------------------------------------------------------- CDP client (as in webauto.cs)

class Cdp : IDisposable
{
    readonly ClientWebSocket socket = new ClientWebSocket();
    readonly List<JsonNode> events = new List<JsonNode>();
    int nextId;

    public static async Task<bool> IsAlive(int port)
    {
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(1) };
            await http.GetStringAsync($"http://127.0.0.1:{port}/json/version");
            return true;
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            return false;
        }
    }

    public static async Task<Cdp> Connect(int port, bool browser = false)
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(3) };
        string url;
        if (browser)
        {
            url = (string)JsonNode.Parse(await http.GetStringAsync($"http://127.0.0.1:{port}/json/version"))["webSocketDebuggerUrl"];
        }
        else
        {
            // Edge opens pages of its own (edge://...), and a browser started on a profile
            // that was used before may not have a tab of ours yet
            JsonNode page = null;
            for (int i = 0; i < 25 && page == null; i++)
            {
                var targets = JsonNode.Parse(await http.GetStringAsync($"http://127.0.0.1:{port}/json")).AsArray();
                page = targets.FirstOrDefault(target => (string)target["type"] == "page" && !((string)target["url"]).StartsWith("edge://"));
                if (page == null)
                {
                    if (i == 10)
                    {
                        using var browserConnection = await Connect(port, browser: true);
                        await browserConnection.Send("Target.createTarget", new JsonObject { ["url"] = "about:blank" });
                    }

                    await Task.Delay(TimeSpan.FromMilliseconds(200));
                }
            }

            url = (string)(page ?? throw new Exception("no page target"))["webSocketDebuggerUrl"];
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
        await socket.SendAsync(Encoding.UTF8.GetBytes(message.ToJsonString()), WebSocketMessageType.Text, endOfMessage: true, CancellationToken.None);
        while (true)
        {
            var reply = await Receive(TimeSpan.FromMinutes(2));
            if (reply["id"] != null && (int)reply["id"] == id)
            {
                if (reply["error"] != null)
                {
                    throw new Exception($"{method}: {reply["error"]["message"]}");
                }

                return reply["result"];
            }

            // network events are many (with --phone) and never waited for
            if ((string)reply["method"] is string name && !name.StartsWith("Network."))
            {
                events.Add(reply);
            }
        }
    }

    public async Task<JsonNode> WaitForEvent(string method, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (true)
        {
            var found = events.FirstOrDefault(received => (string)received["method"] == method);
            if (found != null)
            {
                return found;
            }

            var remaining = deadline - DateTime.UtcNow;
            if (remaining <= TimeSpan.Zero)
            {
                throw new Exception("timed out waiting for " + method);
            }

            events.Add(await Receive(remaining));
        }
    }

    async Task<JsonNode> Receive(TimeSpan timeout)
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
