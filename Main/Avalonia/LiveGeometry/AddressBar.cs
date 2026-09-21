using System;

namespace LiveGeometry;

/// <summary>
/// Where the app is, as a path: "/" is the gallery, "/gallery/morley" a drawing of the gallery,
/// "/drawing" a drawing of the user's own. In the browser this is the real address bar and its
/// history (the Browser head replaces <see cref="Current"/>), so Back works and a link to a
/// drawing can be shared. On the desktop nothing listens, except the window for its title.
/// </summary>
public class AddressBar
{
    public static AddressBar Current { get; set; } = new AddressBar();

    /// <summary>The path the app was started with</summary>
    public virtual string Path => "/";

    /// <summary>The user went back or forward in the history: show what is at this path</summary>
    public event Action<string> PathChanged;

    public event Action<string> TitleChanged;

    /// <summary>The app went somewhere by itself: a new history entry</summary>
    public virtual void Push(string path, string title)
    {
        TitleChanged?.Invoke(title);
    }

    /// <summary>Corrects the current entry (an unknown path that ended up in the gallery)</summary>
    public virtual void Replace(string path, string title)
    {
        TitleChanged?.Invoke(title);
    }

    protected void RaisePathChanged(string path)
    {
        PathChanged?.Invoke(path);
    }
}
