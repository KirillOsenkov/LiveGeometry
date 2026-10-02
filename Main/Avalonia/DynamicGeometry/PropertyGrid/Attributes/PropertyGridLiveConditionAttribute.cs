using System;

namespace DynamicGeometry;

/// <summary>
/// On a button's method of an <see cref="IConditionalProperties"/> object: the button is
/// made also while the object vetoes it, hidden, and shows and hides as the object changes
/// (any change it raises with a property name), without the grid being built again - which
/// would fold a color picker open under the cursor (the drawing's "Reset to default", which
/// appears with the first color picked).
/// </summary>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
public class PropertyGridLiveConditionAttribute : Attribute
{
}
