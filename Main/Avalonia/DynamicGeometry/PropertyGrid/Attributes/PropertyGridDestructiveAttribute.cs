using System;

namespace DynamicGeometry;

/// <summary>
/// Marks a method that deletes something. Its button goes to the very bottom of the
/// property grid, set apart from everything else, and gets a trash can icon.
/// </summary>
[AttributeUsage(AttributeTargets.Method)]
public class PropertyGridDestructiveAttribute : Attribute
{
}
