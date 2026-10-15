using System;
using System.Collections.Generic;
using System.Globalization;
using DynamicGeometry;

namespace LiveGeometry;

/// <summary>
/// The tools the user defined (Define figure), kept between runs: each tool's macro is a
/// document of its own in <see cref="SettingsStore"/> - a file in a Tools folder beside the
/// settings on the desktop, a local storage entry in the browser - stored when the tool is
/// made, again when it is renamed, and deleted with its button. Read at startup, after the
/// ribbon, in the order they were made. Tools and drawings are separate: what a tool makes is
/// ordinary figures, so a drawing never needs the tool.
/// </summary>
public class StoredTools : ToolStorage
{
    const string Group = "Tools";

    // each tool's document, by key: a time stamp (the order they were made) and a random part
    readonly Dictionary<UserDefinedTool, string> keys = new Dictionary<UserDefinedTool, string>();

    // tools being read: adding one to the ribbon is no reason to store it again
    bool loading;

    static SettingsStore Store => SettingsStore.Current;

    /// <summary>Puts the stored tools on the ribbon; one that can't be read is left where it is, with a line on the console</summary>
    public void Load()
    {
        foreach (var key in Store.GetDocumentKeys(Group))
        {
            var tool = UserDefinedTool.Read(Store.GetDocument(Group, key), out string problem);
            if (tool == null)
            {
                Console.WriteLine("Tools: " + key + " is left out. " + problem);
                continue;
            }

            keys[tool] = key;

            // a tool of this version may have taken the name since: the stored one gives way
            if (Behavior.IsToolNameTaken(tool.Name))
            {
                tool.MutableName = Behavior.UniqueToolName(tool.Name);
            }

            loading = true;
            try
            {
                Behavior.Add(tool);
            }
            finally
            {
                loading = false;
            }
        }
    }

    public override void AddTool(UserDefinedTool newBehavior)
    {
        if (loading)
        {
            return;
        }

        var key = DateTime.UtcNow.ToString("yyyyMMdd-HHmmssfff", CultureInfo.InvariantCulture)
            + "-" + Guid.NewGuid().ToString("N").Substring(0, 8);
        keys[newBehavior] = key;
        Save(newBehavior);
    }

    public override void RenameTool(UserDefinedTool behavior, string newName)
    {
        if (keys.ContainsKey(behavior))
        {
            Save(behavior);
        }
    }

    public override void RemoveTool(UserDefinedTool behavior)
    {
        if (keys.Remove(behavior, out var key))
        {
            Store.SetDocument(Group, key, null);
        }
    }

    void Save(UserDefinedTool tool)
    {
        Store.SetDocument(Group, keys[tool], tool.RootElement.ToString());
    }
}
