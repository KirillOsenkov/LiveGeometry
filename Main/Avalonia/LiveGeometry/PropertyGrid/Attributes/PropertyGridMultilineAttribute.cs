using System;

namespace LiveGeometry;

/// <summary>
/// A string property whose editor takes line breaks (Enter starts a new line): the text of
/// a label. Without it a string editor is one line - an expression, a name - and Enter is
/// left to the panel (Plot, Add point). Its box is written larger than the other rows
/// (<see cref="PropertyGridTheme.LabelTextFontSize"/>).
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public class PropertyGridMultilineAttribute : Attribute
{
}
