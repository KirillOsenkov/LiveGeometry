using System.Linq;

namespace DynamicGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("Name")]</c> on a figure's name. Like
/// <see cref="ExpressionEditor"/>, a keystroke that makes a good name applies it, and what is
/// wrong shows under the box on commit: no name at all, one another figure has (files and
/// expressions find figures by name, and the name would be taken from the other figure), or
/// one that reads like a default name the figure would not keep
/// (<see cref="FigureBase.KeepsTypedName"/>).
/// </summary>
public class NameEditorFactory : BaseValueEditorFactory<NameEditor, string>
{
    public NameEditorFactory()
    {
        // after the plain string editor: only the attribute chooses this one
        LoadOrder = 1;
    }
}

public class NameEditor : StringEditor
{
    protected override ValidationResult Validate(object value)
    {
        string name = value.ToString();
        string error = null;
        if (name.Trim().Length == 0)
        {
            error = "A figure needs a name.";
        }
        else if (Value.Parent is IFigure figure
            && figure.Drawing != null
            && figure.Drawing.Figures.Any(f => f != figure && f.Name == name))
        {
            error = "Another figure is already called " + name + ".";
        }
        else if (Value.Parent is FigureBase named && !named.KeepsTypedName(name))
        {
            // AB3 for the segment AB: the name would go straight back to AB
            error = "Names like " + name + " are given automatically. Type another name.";
        }

        return new ValidationResult()
        {
            IsValid = error == null,
            Value = name,
            Error = error
        };
    }
}
