using System.Collections.Generic;

namespace DynamicGeometry;

/// <summary>
/// Something whose properties may hold other values under another theme than the base one
/// (Light): a style's colors, a drawing's paper. The property grid edits what is on screen:
/// under the base theme the property itself, under another theme its override for that
/// theme (<see cref="ThemedValue"/>); a file carries the overrides as a child element per
/// theme (<c>&lt;Dark Color="..." /&gt;</c>).
/// </summary>
public interface IThemeOverridable
{
    /// <summary>By theme name, the properties that differ under that theme, with their values</summary>
    Dictionary<string, Dictionary<string, object>> Overrides { get; }

    void SetOverride(string theme, string property, object value);

    /// <summary>Back to the base theme's values under the named theme</summary>
    void ClearOverrides(string theme);
}
