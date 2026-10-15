namespace DynamicGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("ToolName")]</c> on the name of a tool the user
/// defined. As <see cref="NameEditor"/> does for a figure, a keystroke that makes a good name
/// applies it, and what is wrong shows under the box on commit: no name at all, or one
/// another tool on the ribbon has (two buttons saying the same could not be told apart).
/// </summary>
public class ToolNameEditorFactory : BaseValueEditorFactory<ToolNameEditor, string>
{
    public ToolNameEditorFactory()
    {
        // after NameEditorFactory: "Name", a figure's, is in this one's name too
        LoadOrder = 2;
    }
}

public class ToolNameEditor : StringEditor
{
    protected override ValidationResult Validate(object value)
    {
        string name = value.ToString().Trim();
        string error = null;
        if (name.Length == 0)
        {
            error = "A tool needs a name.";
        }
        else if (Behavior.IsToolNameTaken(name, except: (Value.Parent as UserDefinedTool.UserDefinedDialog)?.Tool))
        {
            error = "Another tool is already called " + name + ".";
        }

        return new ValidationResult()
        {
            IsValid = error == null,
            Value = name,
            Error = error
        };
    }
}
