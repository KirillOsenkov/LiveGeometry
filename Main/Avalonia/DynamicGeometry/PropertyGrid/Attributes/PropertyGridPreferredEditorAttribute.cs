using System;

namespace DynamicGeometry
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Method,
        AllowMultiple = false,
        Inherited = true)]
    public class PropertyGridPreferredEditorAttribute : Attribute
    {
        private readonly string editorTypeName;

        public PropertyGridPreferredEditorAttribute(string editorTypeName)
        {
            this.editorTypeName = editorTypeName;
        }

        public string EditorTypeName
        {
            get
            {
                return this.editorTypeName;
            }
        }
    }
}
