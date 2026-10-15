using System;

namespace DynamicGeometry;

/// <summary>
/// Properties and methods with the same group name are shown together in the property
/// grid, in one box: the editors first, then the buttons in a row.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Method)]
public class PropertyGridGroupAttribute : Attribute
{
    public PropertyGridGroupAttribute(string name)
    {
        Name = name;
    }

    public string Name { get; }
}
