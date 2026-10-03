using System;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices.JavaScript;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Browser;
using LiveGeometry;

internal sealed partial class Program
{
    private static async Task Main(string[] args)
    {
        try
        {
            App.UseInvariantCulture();
            AddressBar.Current = new LiveGeometry.Browser.BrowserAddressBar();
            SettingsStore.Current = new LiveGeometry.Browser.BrowserSettingsStore();

            // hints name the keys as the keyboard does (Option on a Mac)
            DynamicGeometry.KeyNames.IsMac = IsMacKeyboard();

            // a reload, or another visit, finds the user's drawing where it was left
            MainView.KeepsOwnDrawing = true;
            DynamicGeometry.EmojiFont.Open = OpenEmojiFont;
            await BuildAvaloniaApp()
                .WithInterFont()
                .StartBrowserAppAsync("out");
        }
        catch (System.Exception e)
        {
            // Print the exception chain without ex.ToString(): computing stack
            // traces can itself fault on mono-wasm and hide the original error.
            for (var ex = e; ex != null; ex = ex.InnerException)
            {
                System.Console.WriteLine($"CRASH: {ex.GetType().FullName}: {ex.Message}");
            }
        }
    }

    [JSImport("isMac", "main.js")]
    private static partial bool IsMacKeyboard();

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>();

    /// <summary>The emoji font, fetched from the site the first time a character is drawn</summary>
    static async Task<Stream> OpenEmojiFont()
    {
        var document = JSHost.GlobalThis.GetPropertyAsJSObject("document");
        var address = new Uri(new Uri(document.GetPropertyAsString("baseURI")), "fonts/Twemoji.Mozilla.ttf");
        using var client = new HttpClient();
        using var request = new HttpRequestMessage(HttpMethod.Get, address);

        // in one piece: a response streamed in chunks (the default since .NET 10) waits for a
        // turn of the UI thread at every chunk, and while the gallery's tiles were loading the
        // font came in seconds after the tiles that show emoji
        request.Options.Set(new HttpRequestOptionsKey<bool>("WebAssemblyEnableStreamingResponse"), false);
        using var response = await client.SendAsync(request);
        response.EnsureSuccessStatusCode();
        return new MemoryStream(await response.Content.ReadAsByteArrayAsync());
    }
}
