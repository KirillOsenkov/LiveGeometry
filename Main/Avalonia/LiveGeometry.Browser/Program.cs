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

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>();

    /// <summary>The emoji font, fetched from the site the first time a character is drawn</summary>
    static async Task<Stream> OpenEmojiFont()
    {
        var document = JSHost.GlobalThis.GetPropertyAsJSObject("document");
        var address = new Uri(new Uri(document.GetPropertyAsString("baseURI")), "fonts/Twemoji.Mozilla.ttf");
        using var client = new HttpClient();
        return new MemoryStream(await client.GetByteArrayAsync(address));
    }
}
