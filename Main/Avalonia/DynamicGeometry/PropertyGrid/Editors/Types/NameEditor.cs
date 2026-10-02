using System.Linq;

namespace DynamicGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("Name")]</c> on a figure's name. Like
/// <see cref="ExpressionEditor"/>, a keystroke that makes a good name applies it, and what is
/// wrong shows under the box on commit: no name at all, one an expression can't say
/// (<see cref="Scanner.IsName"/>), one another figure has (files and
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
        else if (!Scanner.IsName(name) && Value.Parent is IFigure renamed && IsNamedInExpressions(renamed))
        {
            // Expressions call figures by their names and follow every rename: one they can't
            // say ("my point") left each label naming the figure with an error. A name that
            // is a caption ("Drag me!", as gallery drawings have) is fine for any other figure.
            error = "Expressions name this figure, and can't say " + name + ". A name for them starts with a letter, then letters, digits or primes: A, P2, A'.";
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

    /// <summary>
    /// Whether an expression names the figure (a label's [A.X], the distance [AB]) or a figure
    /// named after it, whose name follows (segment AB in [AB.Length])
    /// </summary>
    static bool IsNamedInExpressions(IFigure figure)
    {
        return figure.Dependents.Any(dependent => HoldsExpressions(dependent)
            || dependent.HasDefaultName && dependent.Dependents.Any(HoldsExpressions));
    }

    /// <summary>A text label, a point by coordinates, a graph...; not a point's name or a measurement, labels without expressions</summary>
    static bool HoldsExpressions(IFigure figure)
    {
        return figure is IRenamableExpressions && (figure is not LabelBase || figure is Label);
    }
}
