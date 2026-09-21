using System.Runtime.InteropServices.JavaScript;

namespace LiveGeometry.Browser;

/// <summary>
/// The address bar and history of the real browser; the JavaScript half is in main.js.
/// </summary>
public partial class BrowserAddressBar : AddressBar
{
    static BrowserAddressBar instance;

    public BrowserAddressBar()
    {
        instance = this;
    }

    public override string Path => GetPath();

    public override void Push(string path, string title)
    {
        PushState(path, title);
        base.Push(path, title);
    }

    public override void Replace(string path, string title)
    {
        ReplaceState(path, title);
        base.Replace(path, title);
    }

    [JSImport("getPath", "main.js")]
    public static partial string GetPath();

    [JSImport("pushState", "main.js")]
    public static partial void PushState(string path, string title);

    [JSImport("replaceState", "main.js")]
    public static partial void ReplaceState(string path, string title);

    /// <summary>Called by main.js on popstate (Back and Forward)</summary>
    [JSExport]
    public static void OnPopState(string path)
    {
        instance?.RaisePathChanged(path);
    }
}
