using System;

namespace LiveGeometry;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Method,
    AllowMultiple = false, 
    Inherited = true)]
public class PropertyGridFocusAttribute : Attribute
{
}
