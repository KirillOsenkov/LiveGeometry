using System;

namespace DynamicGeometry;

/// <summary>
/// On a class: what the grid edits in it is not part of the drawing - the panel of a tool,
/// holding the angle the next line gets - so an edit is not recorded as an undo step.
/// Otherwise typing in such a panel between two constructions is a step of its own, and
/// Ctrl+Z seems to do nothing.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
public class PropertyGridNoUndoAttribute : Attribute
{
    public static bool IsOn(object selection)
    {
        return selection != null && selection.GetType().IsDefined(typeof(PropertyGridNoUndoAttribute), inherit: true);
    }
}
