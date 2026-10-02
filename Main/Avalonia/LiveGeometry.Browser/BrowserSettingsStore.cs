using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.JavaScript;

namespace LiveGeometry.Browser;

/// <summary>
/// The settings in the browser's local storage, one entry per key; the JavaScript half is
/// in main.js (the splash reads the theme from the same entries before the app is up).
/// </summary>
public partial class BrowserSettingsStore : SettingsStore
{
    public override string Get(string key)
    {
        return GetSetting(key);
    }

    public override void Set(string key, string value)
    {
        SetSetting(key, value);
    }

    public override IReadOnlyList<string> GetDocumentKeys(string group)
    {
        var prefix = group + ".";
        var keys = new List<string>();
        foreach (var key in GetSettingKeys(prefix))
        {
            keys.Add(key.Substring(prefix.Length));
        }

        keys.Sort(StringComparer.Ordinal);
        return keys;
    }

    /// <summary>The keys of the settings that start with the prefix, without the app's own prefix</summary>
    [JSImport("getSettingKeys", "main.js")]
    public static partial string[] GetSettingKeys(string prefix);

    [JSImport("getSetting", "main.js")]
    public static partial string GetSetting(string key);

    [JSImport("setSetting", "main.js")]
    public static partial void SetSetting(string key, string value);

    /// <summary>Called by main.js when the page is hidden: on a phone the last sure moment before the tab is gone</summary>
    [JSExport]
    public static void OnPageHidden()
    {
        RaiseLeaving();
    }
}
