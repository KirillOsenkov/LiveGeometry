using System;

namespace DynamicGeometry;

public enum PropertyGridIcon
{
    Pencil,
    Plus
}

/// <summary>
/// Gives the button of a method a small icon in front of its caption
/// (drawn by <see cref="PropertyGridIcons"/>).
/// </summary>
[AttributeUsage(AttributeTargets.Method)]
public class PropertyGridIconAttribute : Attribute
{
    public PropertyGridIconAttribute(PropertyGridIcon icon)
    {
        Icon = icon;
    }

    public PropertyGridIcon Icon { get; }
}
