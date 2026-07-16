using System;

namespace DynamicGeometry
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Method,
        AllowMultiple = false, 
        Inherited = true)]
    public class PropertyGridVisibleAttribute : Attribute
    {
        public PropertyGridVisibleAttribute()
        {
            Visible = true;
        }

        public PropertyGridVisibleAttribute(bool visible)
        {
            Visible = visible;
        }

        public bool Visible { get; set; }
    }

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Method,
        AllowMultiple = false,
        Inherited = true)]
    public class PropertyGridDisallowMultiEditAttribute : Attribute
    {
        public PropertyGridDisallowMultiEditAttribute()
        {
        }
    }

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Method,
        AllowMultiple = false, 
        Inherited = true)]
    public class PropertyGridCustomValueProvider : Attribute
    {
        public PropertyGridCustomValueProvider(Type type)
        {
            CustomValueProviderType = type;
        }

        public Type CustomValueProviderType { get; set; }
    }
}
