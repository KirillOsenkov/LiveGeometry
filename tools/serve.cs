#:property Nullable=disable
#:property PublishAot=false

// serve - minimal static file server for looking at a published browser build locally.
// Separate from webauto.cs on purpose: a running server locks its own executable, which
// would block rebuilding webauto after an edit.
//
//   dotnet run tools/serve.cs -- <dir> [port=5005] [--overlay <prefix>=<folder>]... [--reload]
//   (blocks; run it in the background)
//
// For a `dotnet publish` output pass its wwwroot folder. No compression, no caching: this is
// for functional checks, not for measuring load time (production is IIS with web.config).
//
// Also serves the JavaScript player's dev page: `serve Main` and open
// http://localhost:5005/Player/dev/index.html. A folder without an index.html is served as
// it is (a folder's own URL, ending in /, answers with a JSON list of its entries), and
// every answer carries the CORS header the site sends for the player's files.
//
// To work on a static page of the site against a publish (the embed page needs the player
// and the drawings at their site paths): `--overlay embed=embed` answers everything under
// /embed/ from the repo's embed folder instead of the publish's copy, and `--reload` makes
// every HTML page reload itself when a file in an overlay folder changes (a script put
// before </body> asks /__reload for the folders' change stamp every half second).
//
//   dotnet run tools/serve.cs -- C:\temp\LiveGeometry\player\publish\wwwroot 5006 --overlay embed=embed --reload
//   then open http://localhost:5006/embed/index.html

using System.Net;
using System.Text;
using System.Text.Json;

var positional = new List<string>();
var overlays = new List<(string Prefix, string Folder)>();
bool liveReload = false;
for (int i = 0; i < args.Length; i++)
{
    if (args[i] == "--overlay" && i + 1 < args.Length)
    {
        var pair = args[++i];
        int equals = pair.IndexOf('=');
        if (equals <= 0)
        {
            Console.WriteLine("--overlay wants <prefix>=<folder>");
            return 1;
        }

        overlays.Add((pair.Substring(0, equals).Trim('/'), Path.GetFullPath(pair.Substring(equals + 1))));
    }
    else if (args[i] == "--reload")
    {
        liveReload = true;
    }
    else
    {
        positional.Add(args[i]);
    }
}

if (positional.Count == 0)
{
    Console.WriteLine("usage: serve <dir> [port=5005] [--overlay <prefix>=<folder>]... [--reload]");
    return 1;
}

var root = Path.GetFullPath(positional[0]);
int port = positional.Count > 1 ? int.Parse(positional[1]) : 5005;
bool isApp = File.Exists(Path.Combine(root, "index.html"));
if (!isApp)
{
    Console.WriteLine($"note: {root} has no index.html, served as a plain folder (no SPA fallback)");
}

var mime = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
{
    [".html"] = "text/html", [".js"] = "text/javascript", [".mjs"] = "text/javascript", [".css"] = "text/css",
    [".json"] = "application/json", [".wasm"] = "application/wasm", [".ico"] = "image/x-icon", [".png"] = "image/png",
    [".svg"] = "image/svg+xml", [".map"] = "application/json", [".woff2"] = "font/woff2", [".ttf"] = "font/ttf",
    [".lgf"] = "application/xml", [".txt"] = "text/plain",
};

// the change stamp of the overlay folders, for --reload: any change in any of them moves it
long changeStamp = DateTime.UtcNow.Ticks;
var watchers = new List<FileSystemWatcher>();
foreach (var overlay in overlays)
{
    Console.WriteLine($"overlay: /{overlay.Prefix}/ from {overlay.Folder}");
    if (!liveReload)
    {
        continue;
    }

    var watcher = new FileSystemWatcher(overlay.Folder) { IncludeSubdirectories = true, EnableRaisingEvents = true };
    FileSystemEventHandler changed = (_, _) => Interlocked.Exchange(ref changeStamp, DateTime.UtcNow.Ticks);
    watcher.Changed += changed;
    watcher.Created += changed;
    watcher.Deleted += changed;
    watcher.Renamed += (_, _) => Interlocked.Exchange(ref changeStamp, DateTime.UtcNow.Ticks);
    watchers.Add(watcher);
}

const string reloadScript = "<script>(function () { var seen = null; setInterval(function () { fetch('/__reload', { cache: 'no-store' }).then(function (r) { return r.text(); }).then(function (stamp) { if (seen === null) { seen = stamp; } else if (stamp !== seen) { location.reload(); } }).catch(function () { }); }, 500); })();</script>";

var listener = new HttpListener();
listener.Prefixes.Add($"http://localhost:{port}/");
listener.Start();
Console.WriteLine($"serving {root} at http://localhost:{port}/  (Ctrl+C to stop)");
while (true)
{
    var context = listener.GetContext();
    ThreadPool.QueueUserWorkItem(_ =>
    {
        try
        {
            var relative = Uri.UnescapeDataString(context.Request.Url.AbsolutePath).TrimStart('/');
            context.Response.Headers["Access-Control-Allow-Origin"] = "*";
            context.Response.Headers["Cache-Control"] = "no-cache";

            if (liveReload && relative == "__reload")
            {
                var stamp = Encoding.ASCII.GetBytes(Interlocked.Read(ref changeStamp).ToString());
                context.Response.ContentType = "text/plain";
                context.Response.ContentLength64 = stamp.Length;
                context.Response.OutputStream.Write(stamp);
                return;
            }

            // an overlay answers for its prefix from its own folder; the publish's copy is not consulted
            var baseFolder = root;
            var path = Path.GetFullPath(Path.Combine(root, relative.Length == 0 && isApp ? "index.html" : relative));
            foreach (var overlay in overlays)
            {
                if (relative == overlay.Prefix || relative.StartsWith(overlay.Prefix + "/", StringComparison.OrdinalIgnoreCase))
                {
                    var rest = relative.Length > overlay.Prefix.Length ? relative.Substring(overlay.Prefix.Length + 1) : "";
                    baseFolder = overlay.Folder;
                    path = Path.GetFullPath(Path.Combine(overlay.Folder, rest.Length == 0 ? "index.html" : rest));
                    break;
                }
            }

            // The same fallbacks as web.config. A folder with an index.html is a page of its
            // own (/history, /embed, /web), and an extensionless path under it is a route of
            // that page (/web/morley): the nearest such folder up the path answers. Else a
            // route of the app (/gallery/morley) is not a file, and the app's page answers.
            if (!File.Exists(path) && Path.GetExtension(path).Length == 0)
            {
                var folderIndex = FindFolderIndex(baseFolder, path);
                if (folderIndex != null)
                {
                    path = folderIndex;
                }
                else if (isApp && baseFolder == root)
                {
                    path = Path.Combine(root, "index.html");
                }
                else if (baseFolder != root && File.Exists(Path.Combine(baseFolder, "index.html")))
                {
                    // an overlaid page's own route (/web/morley with --overlay web=web)
                    path = Path.Combine(baseFolder, "index.html");
                }
            }

            if (!path.StartsWith(baseFolder, StringComparison.OrdinalIgnoreCase))
            {
                context.Response.StatusCode = 404;
            }
            else if (Directory.Exists(path) && File.Exists(Path.Combine(path, "index.html")))
            {
                // the page's folder with a trailing slash (/web/) is the page too
                ServeFile(context, Path.Combine(path, "index.html"));
            }
            else if (Directory.Exists(path))
            {
                // a folder as a JSON list of its entries, for a page that wants to know what is there
                var entries = Directory.GetFileSystemEntries(path).Select(Path.GetFileName).Order().ToArray();
                var bytes = JsonSerializer.SerializeToUtf8Bytes(entries);
                context.Response.ContentType = "application/json";
                context.Response.ContentLength64 = bytes.Length;
                context.Response.OutputStream.Write(bytes);
            }
            else if (!File.Exists(path))
            {
                context.Response.StatusCode = 404;
            }
            else
            {
                ServeFile(context, path);
            }
        }
        catch
        {
            // client went away mid-response
        }
        finally
        {
            try { context.Response.Close(); } catch { }
        }
    });
}

void ServeFile(HttpListenerContext context, string path)
{
    var type = mime.TryGetValue(Path.GetExtension(path), out var known) ? known : "application/octet-stream";
    var bytes = File.ReadAllBytes(path);
    if (liveReload && type == "text/html")
    {
        // the page reloads itself when an overlay folder changes
        var text = Encoding.UTF8.GetString(bytes);
        int body = text.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
        text = body >= 0 ? text.Insert(body, reloadScript) : text + reloadScript;
        bytes = Encoding.UTF8.GetBytes(text);
    }

    context.Response.ContentType = type;
    context.Response.ContentLength64 = bytes.Length;
    context.Response.OutputStream.Write(bytes);
}

/// <summary>
/// The index.html of the nearest folder above the path (the base folder itself left out,
/// which is the app's or an overlay's own fallback), or null
/// </summary>
static string FindFolderIndex(string baseFolder, string path)
{
    var folder = Path.GetDirectoryName(path);
    while (folder != null && folder.Length > baseFolder.Length && folder.StartsWith(baseFolder, StringComparison.OrdinalIgnoreCase))
    {
        var index = Path.Combine(folder, "index.html");
        if (File.Exists(index))
        {
            return index;
        }

        folder = Path.GetDirectoryName(folder);
    }

    return null;
}
