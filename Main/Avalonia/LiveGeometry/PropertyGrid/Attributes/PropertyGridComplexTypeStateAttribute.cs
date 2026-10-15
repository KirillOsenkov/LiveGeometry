using System;

namespace LiveGeometry;

public class PropertyGridComplexTypeStateAttribute : Attribute
{
    public PropertyGridComplexTypeStateAttribute(ComplexTypeState state)
    {
        State = state;
    }

    public ComplexTypeState State { get; set; }
}