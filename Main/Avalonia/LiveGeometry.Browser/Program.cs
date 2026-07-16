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
}
