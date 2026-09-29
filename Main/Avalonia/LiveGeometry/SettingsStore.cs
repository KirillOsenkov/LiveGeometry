using System.Collections.Generic;

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
}
