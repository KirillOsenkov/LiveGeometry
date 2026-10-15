using System;
using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>
/// Where the app's own settings live between runs (the theme, the window's place): text
/// values by key. Each head replaces <see cref="Current"/> with what it has - a file in the
/// user's local app data on the desktop, the browser's local storage on the web; this one
/// remembers nothing past the run.
/// </summary>
public class SettingsStore
{
    public static SettingsStore Current { get; set; } = new SettingsStore();

    /// <summary>
    /// The app may be gone in a moment - the window is closing, the browser's tab was hidden
    /// (the last sure moment on a phone): whoever waits to store something stores it now
    /// </summary>
    public static event Action Leaving;

    public static void RaiseLeaving()
    {
        Leaving?.Invoke();
    }

    readonly Dictionary<string, string> values = new Dictionary<string, string>();

    /// <summary>The value stored under the key, or null</summary>
    public virtual string Get(string key)
    {
        return values.TryGetValue(key, out var value) ? value : null;
    }

    /// <summary>Stores the value; null forgets the key</summary>
    public virtual void Set(string key, string value)
    {
        if (value == null)
        {
            values.Remove(key);
        }
        else
        {
            values[key] = value;
        }
    }

    // Documents: longer texts kept by group (the user's tools), each under a key of its own,
    // so that two windows or tabs storing one each don't write over each other's. Here and in
    // the browser they are settings named "group.key"; the desktop keeps a file for each.

    /// <summary>The keys of the documents in the group, in order</summary>
    public virtual IReadOnlyList<string> GetDocumentKeys(string group)
    {
        var prefix = group + ".";
        return values.Keys
            .Where(key => key.StartsWith(prefix, StringComparison.Ordinal))
            .Select(key => key.Substring(prefix.Length))
            .OrderBy(key => key, StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>The document's text, or null</summary>
    public virtual string GetDocument(string group, string key)
    {
        return Get(group + "." + key);
    }

    /// <summary>Stores the document; null deletes it</summary>
    public virtual void SetDocument(string group, string key, string text)
    {
        Set(group + "." + key, text);
    }
}
