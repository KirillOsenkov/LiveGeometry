using System;
using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("Name")]</c> on a figure's name. Like
/// <see cref="ExpressionEditor"/>, a keystroke that makes a good name applies it, and what is
/// wrong shows under the box on commit: no name at all, one another figure has (files and
/// expressions find figures by name, and the name would be taken from the other figure), or
/// one that reads like a default name the figure would not keep
/// (<see cref="FigureBase.KeepsTypedName"/>). Any other name is fine, also for a figure
/// expressions name: they say one that isn't a name of theirs in backticks ("Drag me!" is
/// `Drag me!`, pi is `pi`: <see cref="Scanner.NameInExpression"/>), and the rename writes it so.
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
        string fixCaption = null;
        Action fix = null;
        var figure = Value.Parent as IFigure;
        var other = figure?.Drawing?.Figures.FirstOrDefault(f => f != figure && f.Name == name);
        if (name.Trim().Length == 0)
        {
            error = "A figure needs a name.";
        }
        else if (other != null)
        {
            error = "Another figure is already called " + name + ".";

            // not for a name the figure would not keep anyway (the next check)
            if (figure is not FigureBase named || named.KeepsTypedName(name))
            {
                (fixCaption, fix) = OfferName(figure, other);
            }
        }
        else if (figure is FigureBase named && !named.KeepsTypedName(name))
        {
            // AB3 for the segment AB: the name would go straight back to AB
            error = "Names like " + name + " are given automatically. Type another name.";
        }

        return new ValidationResult()
        {
            IsValid = error == null,
            Value = name,
            Error = error,
            FixCaption = fixCaption,
            Fix = fix
        };
    }

    /// <summary>
    /// The way to the name another figure has: the two swap names (B renamed to A, A is B
    /// now) where the other one keeps the figure's name - not a default name of the
    /// figure's own (segment AB's AB, which would be given to the other figure for no
    /// reason) nor one that reads like the other's default - else the other takes the next
    /// name it would be given
    /// </summary>
    static (string Caption, Action Fix) OfferName(IFigure figure, IFigure other)
    {
        string ownName = figure.Name;
        bool swap = !figure.HasDefaultName && (other is not FigureBase named || named.KeepsTypedName(ownName));
        string othersName = swap ? ownName : other.GenerateFigureName(new List<string>() { other.Name });
        string reference = other is FigureBase figureBase ? figureBase.Reference : other.Name;
        string caption = swap
            ? "Swap names with " + reference
            : "Rename " + reference + " to " + NameDisplay.Format(othersName);
        return (caption, () => Actions.TakeName(figure, other, othersName));
    }
}
