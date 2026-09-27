using System;

namespace DynamicGeometry;

/// <summary>
/// A string property whose editor takes line breaks (Enter starts a new line): the text of
/// a label. Without it a string editor is one line - an expression, a name - and Enter is
/// left to the panel (Plot, Add point).
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public class PropertyGridMultilineAttribute : Attribute
{
}
