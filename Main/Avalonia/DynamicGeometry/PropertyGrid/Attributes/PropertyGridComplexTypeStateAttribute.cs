using System;

namespace DynamicGeometry
{
    public class PropertyGridComplexTypeStateAttribute : Attribute
    {
        public PropertyGridComplexTypeStateAttribute(ComplexTypeState state)
        {
            State = state;
        }

        public ComplexTypeState State { get; set; }
    }
}