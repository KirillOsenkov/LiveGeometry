using System;

namespace DynamicGeometry
{
    [AttributeUsage(AttributeTargets.All, Inherited = true, AllowMultiple = false)]
    public sealed class PropertyGridNameAttribute : Attribute
    {
        public PropertyGridNameAttribute(string name)
        {
            Name = name;
        }

        public string Name { get; set; }
    }
}
