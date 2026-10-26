// Renders the preview picture of every gallery drawing: what a link to /gallery/<slug> shows
// on social sites and in chats (the og:image of the page the WriteGalleryIndex task of
// LiveGeometry.Browser.csproj writes for each drawing).
//
//   dotnet run tools/previews.cs -- <publish>/wwwroot [--browser <path>]
//
// Each drawing of gallery/index.json is played by the JavaScript player (player/1/player.js)
// in a headless Chrome or Edge as its gallery tile shows it - the caption left out, the
// figure fitted - but on its paper, 1200 x 630 (600 x 315 drawn at twice the pixels, so
// that points, lines and names are as big against the picture as on a tile), and written to
// gallery/previews/<File>.png; and Catriona Agg's puzzles that have a drawing
// (CatrionaAgg/puzzles.json) to CatrionaAgg/previews/<slug>.png, with index.png, a mosaic of
// the newest, for the library's page. Emoji are drawn in the site's own font (fonts/), not in
// whatever the system has. CI runs this after the publish
// (.github/workflows/main_livegeometry.yml), on Linux with the Chrome the runner has.
//
// The browser: --browser, else CHROME_PATH or CHROME_BIN, else Chrome, Chromium or Edge where
// they are usually installed. It is started here and closed at the end.

#nullable disable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;

if (args.Length == 0 || args[0].StartsWith("--"))
{
    Console.WriteLine("usage: dotnet run tools/previews.cs -- <publish>/wwwroot [--browser <path>]");
    return 2;
}

string root = Path.GetFullPath(args[0]);
int browserIndex = Array.IndexOf(args, "--browser");
string browser = FindBrowser(browserIndex >= 0 && browserIndex + 1 < args.Length ? args[browserIndex + 1] : null);
string catalogFile = Path.Combine(root, "gallery", "index.json");
string playerFile = Path.Combine(root, "player", "1", "player.js");
string fontFile = Path.Combine(root, "fonts", "Twemoji.Mozilla.ttf");
foreach (var file in new[] { catalogFile, playerFile, fontFile })
{
    if (!File.Exists(file))
    {
        Console.WriteLine($"error: {file} not found: give the wwwroot of a publish of LiveGeometry.Browser");
        return 2;
    }
}

var output = Path.Combine(root, "gallery", "previews");
Directory.CreateDirectory(output);
var profile = Path.Combine(Path.GetTempPath(), "livegeometry-previews-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(profile);
Process process = null;
var watch = Stopwatch.StartNew();
try
{
    process = StartBrowser(browser, profile);
    int port = await WaitForDebugPort(profile, process);
    using var cdp = await Cdp.Connect(port);
    // the 0: the script's last value is not asked to come back
    await cdp.Evaluate(File.ReadAllText(playerFile) + "\n;0");
    await cdp.Evaluate(Setup.Replace("FONT_BASE64", Convert.ToBase64String(File.ReadAllBytes(fontFile))));

    int count = 0;
    var warnings = new List<string>();
    foreach (var item in JsonNode.Parse(File.ReadAllText(catalogFile)).AsArray())
    {
        string file = (string)item["file"];
        string text = File.ReadAllText(Path.Combine(root, "gallery", file));
        var result = JsonNode.Parse(await cdp.Evaluate("renderPreview(" + JsonValue.Create(text).ToJsonString() + ")"));
        string png = (string)result["png"];
        File.WriteAllBytes(Path.Combine(output, Path.ChangeExtension(file, ".png")), Convert.FromBase64String(png.Substring(png.IndexOf(',') + 1)));
        if (result["errors"] is JsonNode errors)
        {
            warnings.Add($"{file}: {errors}");
        }

        count++;
    }

    // Catriona Agg's puzzles (CatrionaAgg/puzzles.json, the ones with a drawing): each one's
    // picture without its handles, as its tile shows it, and a mosaic of the newest for the
    // library's own page
    string catrionaFolder = Path.Combine(root, "CatrionaAgg");
    string catrionaCatalog = Path.Combine(catrionaFolder, "puzzles.json");
    if (File.Exists(catrionaCatalog))
    {
        string catrionaOutput = Path.Combine(catrionaFolder, "previews");
        Directory.CreateDirectory(catrionaOutput);
        var mosaic = new JsonArray();
        foreach (var puzzle in JsonNode.Parse(File.ReadAllText(catrionaCatalog))["puzzles"].AsArray())
        {
            if (puzzle["drawing"] == null)
            {
                continue;
            }

            string slug = (string)puzzle["slug"];
            string text = File.ReadAllText(Path.Combine(catrionaFolder, "drawings", (string)puzzle["drawing"]));
            var result = JsonNode.Parse(await cdp.Evaluate("renderPreview(" + JsonValue.Create(text).ToJsonString() + ", true)"));
            string png = (string)result["png"];
            File.WriteAllBytes(Path.Combine(catrionaOutput, slug + ".png"), Convert.FromBase64String(png.Substring(png.IndexOf(',') + 1)));
            if (result["errors"] is JsonNode errors)
            {
                warnings.Add($"CatrionaAgg/{slug}: {errors}");
            }

            if (mosaic.Count < 8)
            {
                mosaic.Add((JsonNode)JsonValue.Create(text));
            }

            count++;
        }

        string picture = await cdp.Evaluate("renderMosaic(" + mosaic.ToJsonString() + ")");
        File.WriteAllBytes(Path.Combine(catrionaOutput, "index.png"), Convert.FromBase64String(picture.Substring(picture.IndexOf(',') + 1)));
    }

    foreach (var warning in warnings)
    {
        Console.WriteLine("warning: " + warning);
    }

    Console.WriteLine($"{count} previews in {output} ({watch.Elapsed.TotalSeconds:0.0} s)");
    return 0;
}
catch (Exception ex)
{
    Console.WriteLine("error: " + ex.Message);
    return 1;
}
finally
{
    if (process != null && !process.HasExited)
    {
        process.Kill(entireProcessTree: true);
        process.WaitForExit(10000);
    }

    try
    {
        Directory.Delete(profile, recursive: true);
    }
    catch (IOException)
    {
    }
    catch (UnauthorizedAccessException)
    {
    }
}

static string FindBrowser(string given)
{
    if (given != null)
    {
        return File.Exists(given) ? given : throw new Exception(given + " not found");
    }

    var candidates = new List<string>();
    foreach (var variable in new[] { "CHROME_PATH", "CHROME_BIN" })
    {
        if (Environment.GetEnvironmentVariable(variable) is string path && path.Length > 0)
        {
            candidates.Add(path);
        }
    }

    if (OperatingSystem.IsWindows())
    {
        candidates.Add(Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Google\Chrome\Application\chrome.exe"));
        candidates.Add(Environment.ExpandEnvironmentVariables(@"%LocalAppData%\Google\Chrome\Application\chrome.exe"));
        candidates.Add(Environment.ExpandEnvironmentVariables(@"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"));
        candidates.Add(Environment.ExpandEnvironmentVariables(@"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"));
    }
    else if (OperatingSystem.IsMacOS())
    {
        candidates.Add("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome");
        candidates.Add("/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge");
    }
    else
    {
        var folders = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(':', StringSplitOptions.RemoveEmptyEntries);
        foreach (var name in new[] { "google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "microsoft-edge" })
        {
            candidates.AddRange(folders.Select(folder => Path.Combine(folder, name)));
        }
    }

    return candidates.FirstOrDefault(File.Exists) ?? throw new Exception("no Chrome, Chromium or Edge found: give one with --browser");
}

static Process StartBrowser(string browser, string profile)
{
    // not through the shell: the browser doesn't outlive this tool (killed in the finally),
    // and its chatter on stdout and stderr is read and dropped here instead of in the log
    var info = new ProcessStartInfo(browser)
    {
        UseShellExecute = false,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        CreateNoWindow = true,
    };
    foreach (var argument in new[]
    {
        "--headless=new", "--remote-debugging-port=0", $"--user-data-dir={profile}",
        "--no-first-run", "--no-default-browser-check", "--disable-extensions", "--disable-sync",
        "--hide-scrollbars", "--mute-audio", "--window-size=800,600",
    })
    {
        info.ArgumentList.Add(argument);
    }

    if (Path.GetFileNameWithoutExtension(browser).Equals("msedge", StringComparison.OrdinalIgnoreCase))
    {
        // otherwise Edge signs the throwaway profile into the Windows account and opens pages of its own
        info.ArgumentList.Add("--guest");
    }

    if (OperatingSystem.IsLinux())
    {
        // a CI runner may not allow the user namespaces Chrome's sandbox needs; the page is our own
        info.ArgumentList.Add("--no-sandbox");
    }

    info.ArgumentList.Add("about:blank");
    var process = Process.Start(info);
    process.OutputDataReceived += (sender, e) => { };
    process.ErrorDataReceived += (sender, e) => { };
    process.BeginOutputReadLine();
    process.BeginErrorReadLine();
    return process;
}

// With port 0 the browser picks a free port and writes it into DevToolsActivePort in its profile.
static async Task<int> WaitForDebugPort(string profile, Process process)
{
    var file = Path.Combine(profile, "DevToolsActivePort");
    var deadline = DateTime.UtcNow.AddSeconds(30);
    while (DateTime.UtcNow < deadline)
    {
        if (process.HasExited)
        {
            throw new Exception($"the browser exited at once (exit code {process.ExitCode})");
        }

        try
        {
            if (File.Exists(file) && int.TryParse(File.ReadLines(file).FirstOrDefault(), out int port) && port > 0)
            {
                return port;
            }
        }
        catch (IOException)
        {
            // still being written
        }

        await Task.Delay(100);
    }

    throw new Exception("the browser didn't open its debugging port within 30 s");
}

partial class Program
{
    // Runs in the blank page once the player is loaded: the emoji font, and renderPreview,
    // which plays one drawing and gives its picture as a data URL. Sizes and the fit are the
    // gallery tile's (TilePlayer in web/index.html, DrawingThumbnail in the app).
    const string Setup = """
        (async () => {
            // only for emoji: digits, letters and signs like − come from the text font
            const bytes = Uint8Array.from(atob("FONT_BASE64"), c => c.charCodeAt(0));
            const face = new FontFace("Live Geometry Emoji", bytes, { unicodeRange: "U+00A9, U+00AE, U+203C-3299, U+FE0F, U+200D, U+1F000-1FFFF" });
            await face.load();
            document.fonts.add(face);
            const font = "'Live Geometry Emoji', 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Liberation Sans', sans-serif";
            const Width = 600;
            const Height = 315;
            const Scale = 2;
            const ZoomAfterFit = 1.08;
            document.body.style.margin = "0";

            // 600 x 315, or the size given (set after the base constructor, which measures first)
            class PreviewPlayer extends LiveGeometry.Player {
                measureElement() {
                    const [width, height] = this.size ?? [Width, Height];
                    if (this.canvas.width === width && this.canvas.height === height) {
                        return false;
                    }

                    this.canvas.width = width;
                    this.canvas.height = height;
                    this.pixelRatio = Scale;
                    this.canvasElement.width = width * Scale;
                    this.canvasElement.height = height * Scale;
                    return true;
                }

                show(drawing) {
                    for (const figure of drawing.figures.list) {
                        if (figure instanceof LiveGeometry.Label && figure.pin !== LiveGeometry.LabelPin.None) {
                            figure.visible = false;
                        }
                    }

                    drawing.paintsPaper = true;
                    super.show(drawing);
                }

                fitInitial() {
                    const drawing = this.drawing;
                    if (drawing == null) {
                        return;
                    }

                    const scene = drawing.chooseScene(this.canvas.width, this.canvas.height);
                    if (scene != null) {
                        drawing.showScene(scene);
                    } else {
                        drawing.coordinateSystem.zoomExtend(LiveGeometry.GalleryDrawing.getPlane(drawing), ZoomAfterFit);
                    }
                }
            }

            // the pinned labels nothing is built on (the title and the explanation) are not
            // read at all, as on a tile (leaveOutCaption in web/index.html)
            function leaveOutCaption(element) {
                const figures = [...element.children].find(child => child.localName === "Figures");
                if (figures == null) {
                    return element;
                }

                const named = new Set();
                figures.querySelectorAll("Dependency").forEach(dependency => {
                    if (dependency.getAttribute("Name") != null) {
                        named.add(dependency.getAttribute("Name"));
                    }
                });
                const pinned = [...figures.children].filter(figure => {
                    const pin = figure.getAttribute("Pin");
                    return pin != null && pin !== "None" && !named.has(figure.getAttribute("Name") || "");
                });
                pinned.forEach(figure => figure.remove());
                return element;
            }

            // the points a reader drags are left out of a picture of a Catriona Agg puzzle, as
            // on its tile (withoutHandles in CatrionaAgg/catriona.js)
            function parse(text, hideHandles) {
                const parsed = new DOMParser().parseFromString(text.replace(/^﻿/, "").trimStart(), "application/xml");
                if (parsed.querySelector("parsererror") != null) {
                    throw new Error("not well-formed XML");
                }

                if (hideHandles) {
                    parsed.querySelectorAll('[Style="Handle"]').forEach(figure => figure.setAttribute("Visible", "false"));
                }

                return leaveOutCaption(parsed.documentElement);
            }

            function play(element, use, size) {
                const host = document.createElement("div");
                host.style.width = Width + "px";
                host.style.height = Height + "px";
                document.body.appendChild(host);
                const player = new PreviewPlayer(host, { theme: "light", font: font });
                if (size != null) {
                    player.size = size;
                    player.measureElement();
                }

                try {
                    player.load(element);
                    player.render();
                    return use(player);
                } finally {
                    // the frame the player asked for on loading comes after this
                    player.render = () => { };
                    player.dispose();
                    host.remove();
                }
            }

            window.renderPreview = (text, hideHandles) => play(parse(text, hideHandles), player =>
                JSON.stringify({ png: player.canvasElement.toDataURL("image/png"), errors: player.loadErrors ?? null }));

            // The library's own picture: the drawings as cards on its warm paper, four by two
            window.renderMosaic = texts => {
                const canvas = document.createElement("canvas");
                canvas.width = 1200;
                canvas.height = 630;
                const context = canvas.getContext("2d");
                context.fillStyle = "#F7F5F0";
                context.fillRect(0, 0, canvas.width, canvas.height);
                const columns = 4, gap = 24, card = (canvas.width - gap * (columns + 1)) / columns;
                const top = (canvas.height - 2 * card - gap) / 2;
                texts.forEach((text, index) => {
                    const x = gap + (index % columns) * (card + gap);
                    const y = top + Math.floor(index / columns) * (card + gap);
                    context.save();
                    context.shadowColor = "rgba(40, 32, 20, 0.14)";
                    context.shadowBlur = 18;
                    context.shadowOffsetY = 4;
                    context.fillStyle = "#FFFFFF";
                    context.beginPath();
                    context.roundRect(x, y, card, card, 16);
                    context.fill();
                    context.restore();
                    play(parse(text, true), player => context.drawImage(player.canvasElement, x + 8, y + 8, card - 16, card - 16), [card - 16, card - 16]);
                });
                return canvas.toDataURL("image/png");
            };
        })()
        """;
}

class Cdp : IDisposable
{
    private readonly ClientWebSocket socket = new ClientWebSocket();
    private int nextId;

    public static async Task<Cdp> Connect(int port)
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
        var targets = JsonNode.Parse(await http.GetStringAsync($"http://127.0.0.1:{port}/json/list")).AsArray();
        var page = targets.FirstOrDefault(t => (string)t["type"] == "page" && (string)t["url"] == "about:blank")
            ?? targets.FirstOrDefault(t => (string)t["type"] == "page")
            ?? throw new Exception("the browser has no page");
        var cdp = new Cdp();
        cdp.socket.Options.KeepAliveInterval = TimeSpan.Zero;
        await cdp.socket.ConnectAsync(new Uri((string)page["webSocketDebuggerUrl"]), CancellationToken.None);
        return cdp;
    }

    /// <summary>Runs a script in the page, waits for the promise it gives if any, and returns its value as text</summary>
    public async Task<string> Evaluate(string expression)
    {
        var result = await Send("Runtime.evaluate", new JsonObject { ["expression"] = expression, ["returnByValue"] = true, ["awaitPromise"] = true });
        if (result["exceptionDetails"] is JsonNode details)
        {
            throw new Exception("in the page: " + (details["exception"]?["description"] ?? details["text"]));
        }

        var value = result["result"]?["value"];
        return value == null ? null : value.GetValueKind() == JsonValueKind.String ? (string)value : value.ToJsonString();
    }

    private async Task<JsonNode> Send(string method, JsonObject parameters)
    {
        int id = ++nextId;
        var message = new JsonObject { ["id"] = id, ["method"] = method, ["params"] = parameters };
        await socket.SendAsync(Encoding.UTF8.GetBytes(message.ToJsonString()), WebSocketMessageType.Text, true, CancellationToken.None);
        while (true)
        {
            var reply = await Receive(TimeSpan.FromSeconds(60));
            if (reply["id"] != null && (int)reply["id"] == id)
            {
                return reply["error"] != null ? throw new Exception($"{method}: {reply["error"]["message"]}") : reply["result"];
            }
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
