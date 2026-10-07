using System;
using System.IO;
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
            SplashScreen.Attach(HideSplash, ReportSplashProgress);
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

    // the page's splash screen (index.html), see SplashScreen
    [JSImport("hideSplash", "main.js")]
    private static partial void HideSplash();

    [JSImport("reportProgress", "main.js")]
    private static partial void ReportSplashProgress(double fraction);

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>();

    /// <summary>
    /// The emoji font, fetched from the site when the gallery is built or the first character
    /// is drawn. The page fetches it (main.js) and the app takes it once it is all there: one
    /// turn of the UI thread, where HttpClient took one for every step of the download.
    /// </summary>
    static async Task<Stream> OpenEmojiFont()
    {
        var bytes = new byte[await FetchEmojiFont()];
        CopyEmojiFont(bytes);
        return new MemoryStream(bytes);
    }

    [JSImport("fetchEmojiFont", "main.js")]
    [return: JSMarshalAs<JSType.Promise<JSType.Number>>]
    private static partial Task<int> FetchEmojiFont();

    [JSImport("copyEmojiFont", "main.js")]
    private static partial void CopyEmojiFont([JSMarshalAs<JSType.MemoryView>] Span<byte> target);
}
