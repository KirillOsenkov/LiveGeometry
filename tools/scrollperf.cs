#:property Nullable=disable
#:property PublishAot=false

// scrollperf - how the gallery scrolls: headless Edge over the Chrome DevTools Protocol (port 9335,
// apart from webauto's and loadperf's). The page is loaded and left alone until its tiles are in,
// then scrolled - with a finger (touch events, 60 a second, as a phone sends them) or the mouse
// wheel - twice: down, which loads the tiles it reaches as it goes, and back up over tiles that
// are all loaded, which is the scrolling alone. For each the page records its frames: the time
// between the browser's animation frames (what the eye sees), how long the main thread spent in
// the callbacks of timers and animation frames (Avalonia's dispatcher and render pass run in
// those), and the long tasks. Reports counts and percentiles; the raw records go to down.json
// and up.json in the out folder, screenshots to before.png, down.png and up.png.
//
//   dotnet run tools/scrollperf.cs -- <url> <out folder> [--phone] [--wheel] [--size 1700x1000] [--wait 15] [--passes 3] [--dpr 1]
//
// --phone   a phone: 390x645, the CPU 4 times slower (DevTools' preset), a finger scrolls
// --dpr     the phone's device pixel ratio; 1 by default, not a phone's 3: under headless Edge's
//           emulation Avalonia measures its canvas in physical pixels and divides by the emulated
//           ratio, and laid the page out at a third of its width. The main thread's work is the
//           same at any ratio; only the GPU draws fewer pixels.
// --wheel   the mouse wheel scrolls (the default without --phone)
// --wait    how long the page gets to load before the scroll begins (seconds); it begins earlier
//           once the splash is gone and the main thread has been quiet for two seconds
// --passes  how many scrolls, each down by most of a screen, with a pause for the inertia to settle
//
// Nothing else should run meanwhile: a browser busy on the same machine takes CPU from the page.

using System.Diagnostics;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

const int Port = 9335;

// in the page before its own scripts run: every callback of requestAnimationFrame, setTimeout and
// setInterval is timed (Avalonia's dispatcher runs its jobs and the render pass from those), long
// tasks are watched, and start/stop record the frames of a scroll
const string Hook = """
(() => {
  const record = window.__scrollperf = { frames: [], callbacks: [], longTasks: [], recording: false, skip: false, lastLongTaskEnd: 0 };
  const timed = callback => function () {
    const started = performance.now();
    try {
      return callback.apply(this, arguments);
    } finally {
      if (record.skip) {
        record.skip = false;
      } else if (record.recording) {
        record.callbacks.push([Math.round(started), Math.round((performance.now() - started) * 10) / 10]);
      }
    }
  };
  const raf = window.requestAnimationFrame.bind(window);
  window.requestAnimationFrame = callback => raf(timed(callback));
  const timeout = window.setTimeout.bind(window);
  window.setTimeout = (callback, delay, ...rest) => timeout(typeof callback === 'function' ? timed(callback) : callback, delay, ...rest);
  const interval = window.setInterval.bind(window);
  window.setInterval = (callback, delay, ...rest) => interval(typeof callback === 'function' ? timed(callback) : callback, delay, ...rest);
  try {
    new PerformanceObserver(list => {
      for (const task of list.getEntries()) {
        record.lastLongTaskEnd = task.startTime + task.duration;
        if (record.recording) record.longTasks.push([Math.round(task.startTime), Math.round(task.duration)]);
      }
    }).observe({ type: 'longtask', buffered: true });
  } catch (e) { }
  record.start = () => {
    record.frames = []; record.callbacks = []; record.longTasks = []; record.recording = true;
    const loop = time => { if (!record.recording) return; record.frames.push(Math.round(time * 10) / 10); record.skip = true; raf(loop); };
    record.skip = true;
    raf(loop);
  };
  record.stop = () => {
    record.recording = false;
    return JSON.stringify({ frames: record.frames, callbacks: record.callbacks, longTasks: record.longTasks });
  };
  record.isReady = () => (document.querySelector('.avalonia-splash') === null || document.querySelector('.avalonia-splash.splash-close') !== null)
    && performance.now() - record.lastLongTaskEnd > 2000;
})();
""";

if (args.Length < 2)
{
    Console.WriteLine("usage: scrollperf <url> <out folder> [--phone] [--wheel] [--size 1700x1000] [--wait 15] [--passes 3]");
    return 1;
}

var url = args[0];
var outputFolder = Path.GetFullPath(args[1]);
bool phone = args.Contains("--phone");
bool wheel = args.Contains("--wheel") || !phone;
var size = (ArgumentValue("--size") ?? "1700x1000").Split('x').Select(int.Parse).ToArray();
var wait = TimeSpan.FromSeconds(int.Parse(ArgumentValue("--wait") ?? "15"));
int passes = int.Parse(ArgumentValue("--passes") ?? "3");
int pixelRatio = int.Parse(ArgumentValue("--dpr") ?? "1");
Directory.CreateDirectory(outputFolder);
var profile = Path.Combine(Path.GetTempPath(), "scrollperf-" + Guid.NewGuid().ToString("N"));

await CloseBrowser();
try
{
    await StartBrowser();
    using var page = await Cdp.Connect(Port);
    await page.Send("Page.enable");
    int width = size[0];
    int height = size[1];
    if (phone)
    {
        width = 390;
        height = 645;
        await page.Send("Emulation.setDeviceMetricsOverride", new JsonObject
        {
            ["width"] = width,
            ["height"] = height,
            ["deviceScaleFactor"] = pixelRatio,
            ["mobile"] = true,
        });
        await page.Send("Emulation.setCPUThrottlingRate", new JsonObject { ["rate"] = 4 });
    }

    await page.Send("Page.addScriptToEvaluateOnNewDocument", new JsonObject { ["source"] = Hook });
    await page.Send("Page.navigate", new JsonObject { ["url"] = url });

    // the page loads, and its tiles after it
    var loading = Stopwatch.StartNew();
    while (loading.Elapsed < wait)
    {
        await Task.Delay(500);
        if (loading.Elapsed > TimeSpan.FromSeconds(3) && await Evaluate(page, "window.__scrollperf.isReady()") == "true")
        {
            break;
        }
    }

    // what the page is laid out at (the emulation has caught Avalonia out before, see --dpr)
    Console.WriteLine("page: " + await Evaluate(page, "innerWidth + 'x' + innerHeight + ' at ' + devicePixelRatio + 'x, canvas ' + (document.querySelector('canvas') ? document.querySelector('canvas').width + 'x' + document.querySelector('canvas').height : 'none')"));
    var before = await page.Send("Page.captureScreenshot", new JsonObject { ["format"] = "png" });
    File.WriteAllBytes(Path.Combine(outputFolder, "before.png"), Convert.FromBase64String((string)before["data"]));
    Console.WriteLine($"scrolling after {loading.Elapsed.TotalSeconds:0.0} s of loading ({(phone ? "phone" : "desktop")}, {(wheel ? "the wheel" : "a finger")}, {passes} passes each way)");
    await Measure(page, "down", down: true, width, height);

    // the tiles the way down reached are loading still
    await Task.Delay(3000);
    await Measure(page, "up", down: false, width, height);
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

async Task Measure(Cdp page, string name, bool down, int width, int height)
{
    await Evaluate(page, "window.__scrollperf.start()");
    var scrolling = Stopwatch.StartNew();
    for (int pass = 0; pass < passes; pass++)
    {
        if (wheel)
        {
            // a notch every 50 ms for a second
            for (int i = 0; i < 20; i++)
            {
                await page.Send("Input.dispatchMouseEvent", new JsonObject
                {
                    ["type"] = "mouseWheel", ["x"] = width / 2, ["y"] = height / 2, ["deltaX"] = 0, ["deltaY"] = down ? 120 : -120,
                });
                await Task.Delay(50);
            }

            await Task.Delay(1000);
        }
        else
        {
            // a finger from three quarters down the screen to a fifth (the other way up), in 50
            // steps over 0.8 s
            double x = width / 2.0;
            double from = down ? height * 0.75 : height * 0.2;
            double to = down ? height * 0.2 : height * 0.75;
            const int steps = 50;
            await Touch(page, "touchStart", (x, from));
            for (int i = 1; i <= steps; i++)
            {
                await Touch(page, "touchMove", (x, from + (to - from) * i / steps));
                await Task.Delay(16);
            }

            await Touch(page, "touchEnd");
            await Task.Delay(1500);
        }
    }

    var json = await Evaluate(page, "window.__scrollperf.stop()");
    double scrolled = scrolling.Elapsed.TotalMilliseconds;
    File.WriteAllText(Path.Combine(outputFolder, name + ".json"), json);
    var shot = await page.Send("Page.captureScreenshot", new JsonObject { ["format"] = "png" });
    File.WriteAllBytes(Path.Combine(outputFolder, name + ".png"), Convert.FromBase64String((string)shot["data"]));
    Console.WriteLine();
    Console.WriteLine($"== {name} ==");
    Report(json, scrolled);
}

string ArgumentValue(string name)
{
    int index = Array.IndexOf(args, name);
    return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
}

void Report(string json, double scrolledMilliseconds)
{
    var root = JsonNode.Parse(json);
    var frames = root["frames"].AsArray().Select(frame => (double)frame).ToList();
    var intervals = new List<double>();
    for (int i = 1; i < frames.Count; i++)
    {
        intervals.Add(frames[i] - frames[i - 1]);
    }

    intervals.Sort();
    var callbacks = root["callbacks"].AsArray().Select(callback => (double)callback[1]).OrderBy(duration => duration).ToList();
    var tasks = root["longTasks"].AsArray().Select(task => (Start: (int)task[0], Duration: (int)task[1])).ToList();
    Console.WriteLine($"{scrolledMilliseconds / 1000:0.0} s of scrolling and settling, {frames.Count} frames ({frames.Count / (scrolledMilliseconds / 1000):0.0} a second)");
    if (intervals.Count > 0)
    {
        Console.WriteLine($"frame interval: median {Percentile(intervals, 50):0.0} ms, p90 {Percentile(intervals, 90):0.0}, p95 {Percentile(intervals, 95):0.0}, longest {intervals[^1]:0.0}");
        Console.WriteLine($"  frames over 34 ms (under 30 fps): {intervals.Count(interval => interval > 34)}, over 100 ms: {intervals.Count(interval => interval > 100)}");
    }

    if (callbacks.Count > 0)
    {
        Console.WriteLine($"timer and frame callbacks: {callbacks.Count}, {callbacks.Sum():0} ms in all ({100 * callbacks.Sum() / scrolledMilliseconds:0}% of the time), median {Percentile(callbacks, 50):0.0} ms, p95 {Percentile(callbacks, 95):0.0}, longest {callbacks[^1]:0.0}");
    }

    Console.WriteLine($"long tasks: {tasks.Count}, {tasks.Sum(task => task.Duration)} ms in all" + (tasks.Count > 0 ? ", longest " + tasks.Max(task => task.Duration) + " ms" : ""));
}

static double Percentile(List<double> sorted, int percent)
{
    if (sorted.Count == 0)
    {
        return 0;
    }

    int index = (int)Math.Round((sorted.Count - 1) * percent / 100.0);
    return sorted[Math.Clamp(index, 0, sorted.Count - 1)];
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

// The fingers on the screen after the event (the place in the list is the finger's id); none for
// the end of the last touch.
static Task<JsonNode> Touch(Cdp page, string type, params (double X, double Y)[] fingers)
{
    var points = new JsonArray();
    for (int i = 0; i < fingers.Length; i++)
    {
        points.Add(new JsonObject { ["x"] = fingers[i].X, ["y"] = fingers[i].Y, ["id"] = i, ["radiusX"] = 8, ["radiusY"] = 8, ["force"] = 1 });
    }

    return page.Send("Input.dispatchTouchEvent", new JsonObject { ["type"] = type, ["touchPoints"] = points });
}

static async Task<string> Evaluate(Cdp page, string expression)
{
    var result = await page.Send("Runtime.evaluate", new JsonObject { ["expression"] = expression, ["returnByValue"] = true });
    if (result["exceptionDetails"] != null)
    {
        throw new Exception("js: " + result["exceptionDetails"].ToJsonString());
    }

    var value = result["result"]?["value"];
    return value == null ? "undefined" : value.GetValueKind() == JsonValueKind.String ? (string)value : value.ToJsonString();
}

// ---------------------------------------------------------------- CDP client (as in loadperf.cs)

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

            events.Add(reply);
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
