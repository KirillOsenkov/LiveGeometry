using System;
using System.Collections.Generic;
using System.IO;

namespace LiveGeometry.Desktop;

/// <summary>
/// The settings in a text file of <c>key=value</c> lines in the user's local app data,
/// read once at startup and rewritten whole on every change (it is a handful of lines).
/// </summary>
public class FileSettingsStore : SettingsStore
{
    readonly Dictionary<string, string> values = new Dictionary<string, string>();

    public static string SettingsFile
    {
        get
        {
            var folder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(folder, "LiveGeometry", "Settings.txt");
        }
    }

    public FileSettingsStore()
    {
        try
        {
            if (File.Exists(SettingsFile))
            {
                foreach (var line in File.ReadAllLines(SettingsFile))
                {
                    int separator = line.IndexOf('=');
                    if (separator > 0)
                    {
                        values[line.Substring(0, separator).Trim()] = line.Substring(separator + 1).Trim();
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // never a reason not to start
            Console.WriteLine("Settings: " + ex.Message);
        }
    }

    public override string Get(string key)
    {
        return values.TryGetValue(key, out var value) ? value : null;
    }

    public override void Set(string key, string value)
    {
        if (value == null)
        {
            values.Remove(key);
        }
        else
        {
            values[key] = value;
        }

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsFile));
            var lines = new List<string>();
            foreach (var pair in values)
            {
                lines.Add(pair.Key + "=" + pair.Value);
            }

            File.WriteAllLines(SettingsFile, lines);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Settings: " + ex.Message);
        }
    }

    /// <summary>A group's documents are files in a folder of its name beside the settings file: Tools\*.xml</summary>
    static string GetFolder(string group)
    {
        return Path.Combine(Path.GetDirectoryName(SettingsFile), group);
    }

    public override IReadOnlyList<string> GetDocumentKeys(string group)
    {
        var keys = new List<string>();
        var folder = GetFolder(group);
        if (!Directory.Exists(folder))
        {
            return keys;
        }

        try
        {
            foreach (var file in Directory.GetFiles(folder, "*.xml"))
            {
                keys.Add(Path.GetFileNameWithoutExtension(file));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Settings: " + ex.Message);
        }

        keys.Sort(StringComparer.Ordinal);
        return keys;
    }

    public override string GetDocument(string group, string key)
    {
        var file = Path.Combine(GetFolder(group), key + ".xml");
        try
        {
            return File.Exists(file) ? File.ReadAllText(file) : null;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Settings: " + ex.Message);
            return null;
        }
    }

    public override void SetDocument(string group, string key, string text)
    {
        var file = Path.Combine(GetFolder(group), key + ".xml");
        try
        {
            if (text == null)
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }

                return;
            }

            Directory.CreateDirectory(GetFolder(group));
            File.WriteAllText(file, text);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Settings: " + ex.Message);
        }
    }
}
