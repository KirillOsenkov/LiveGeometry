using System;

namespace LiveGeometry;

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

    public string EditorTypeName => this.editorTypeName;
}
