using System;

namespace LiveGeometry;

/// <summary>
/// The parts of a page of the property grid, top to bottom, as a figure's context menu has
/// them (<see cref="ContextMenus"/>)
/// </summary>
public enum PropertyGridSection
{
    /// <summary>Over the rows, staying put while they scroll: the name, and Visible and Locked as toggles</summary>
    Header,
    /// <summary>What the figure is: its coordinates, length, radius, text, formula (rows unmarked)</summary>
    Definition,
    /// <summary>How it shows itself: Show name, a segment's mark, an angle's arcs, decimals</summary>
    Display,
    /// <summary>The style, with Edit this style and Create new style</summary>
    Style,
    /// <summary>What it can do: Free point, Convert to (buttons unmarked)</summary>
    Verbs,
    /// <summary>Bring to front, Send to back</summary>
    Arrange,
    /// <summary>Under a divider, last: Delete (<see cref="PropertyGridDestructiveAttribute"/>)</summary>
    Delete
}

/// <summary>
/// Where a row or a button goes on its page: its section, and its place in the section,
/// smaller first (rows and buttons of the same place stay in the order they are declared,
/// a class's own before its base's). Unmarked, a row is <see cref="PropertyGridSection.Definition"/>,
/// a button <see cref="PropertyGridSection.Verbs"/> and a destructive button
/// <see cref="PropertyGridSection.Delete"/>, all at 0. Places go in tens, so that a class
/// can put a row of its own between two of its base's. A group box
/// (<see cref="PropertyGridGroupAttribute"/>) goes where the first of its members goes.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Method, Inherited = true, AllowMultiple = false)]
public class PropertyGridSectionAttribute : Attribute
{
    public PropertyGridSectionAttribute(PropertyGridSection section, int order = 0)
    {
        Section = section;
        Order = order;
    }

    public PropertyGridSection Section { get; }

    public int Order { get; }
}
