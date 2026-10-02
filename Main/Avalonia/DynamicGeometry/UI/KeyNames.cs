using System;

namespace DynamicGeometry;

/// <summary>
/// Keys as the user's keyboard names them, for the text of hints: Alt is Option on a Mac.
/// A class of its own, with nothing else static: the browser sets <see cref="IsMac"/> before
/// Avalonia is up, and a class that also makes cursors or brushes in its static constructor
/// (as <see cref="Behavior"/> does) can't be touched then - the app died before its first frame.
/// </summary>
public static class KeyNames
{
    /// <summary>
    /// Whether the keyboard is a Mac's (an iPad's too). The desktop knows; the browser head
    /// sets it from the page's platform, since a browser says it runs on "browser".
    /// </summary>
    public static bool IsMac { get; set; } = OperatingSystem.IsMacOS();

    public static string Alt => IsMac ? "Option" : "Alt";
}
