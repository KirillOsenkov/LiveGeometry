#:property Nullable=disable
#:property PublishAot=false

// serve - minimal static file server for looking at a published browser build locally.
// Separate from webauto.cs on purpose: a running server locks its own executable, which
// would block rebuilding webauto after an edit.
//
//   dotnet run tools/serve.cs -- <dir> [port=5005]      (blocks; run it in the background)
//
// For a `dotnet publish` output pass its wwwroot folder. No compression, no caching: this is
// for functional checks, not for measuring load time (production is IIS with web.config).

using System.Net;

if (args.Length == 0)
{
    Console.WriteLine("usage: serve <dir> [port=5005]");
    return 1;
}

var root = Path.GetFullPath(args[0]);
int port = args.Length > 1 ? int.Parse(args[1]) : 5005;
if (!File.Exists(Path.Combine(root, "index.html")))
{
    Console.WriteLine($"error: {root} has no index.html (for a publish output pass its wwwroot folder)");
    return 2;
}

var mime = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
{
    [".html"] = "text/html", [".js"] = "text/javascript", [".mjs"] = "text/javascript", [".css"] = "text/css",
    [".json"] = "application/json", [".wasm"] = "application/wasm", [".ico"] = "image/x-icon", [".png"] = "image/png",
    [".svg"] = "image/svg+xml", [".map"] = "application/json", [".woff2"] = "font/woff2",
};

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
            var path = Path.GetFullPath(Path.Combine(root, relative.Length == 0 ? "index.html" : relative));

            // the same fallback as web.config: a route of the app (/gallery/morley) is not a file
            if (!File.Exists(path) && Path.GetExtension(path).Length == 0)
            {
                path = Path.Combine(root, "index.html");
            }

            if (!path.StartsWith(root, StringComparison.OrdinalIgnoreCase) || !File.Exists(path))
            {
                context.Response.StatusCode = 404;
            }
            else
            {
                context.Response.ContentType = mime.TryGetValue(Path.GetExtension(path), out var type) ? type : "application/octet-stream";
                context.Response.Headers["Cache-Control"] = "no-cache";
                var bytes = File.ReadAllBytes(path);
                context.Response.ContentLength64 = bytes.Length;
                context.Response.OutputStream.Write(bytes);
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
