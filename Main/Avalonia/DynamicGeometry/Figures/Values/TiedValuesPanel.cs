using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry;

/// <summary>
/// The side panel right after a figure with tied values is made (a line at an angle, a
/// rotated or dilated figure, a translated one): the figure's own value rows - editable
/// while typed, naming the source while tied - a "Type the ..." button for each tied one, and
/// OK, which brings the tool's panel back. While it is up, a click on a suitable figure ties
/// a value to it (<see cref="FigureCreator.TryTieCreatedFigure"/>). The rows are the
/// figure's (its own grid has the same), so an edit is an undo step on the figure. The title
/// is the figure made ("Triangle A'B'C'"); the rows come from one of its points.
/// </summary>
public class TiedValuesPanel : ICustomPropertyProvider, ICustomMethodProvider
{
    public TiedValuesPanel(IFigure shown, ITiedValues values)
    {
        Shown = shown;
        Values = values;
    }

    public IFigure Shown { get; }

    public ITiedValues Values { get; }

    public IEnumerable<IValueProvider> GetProperties()
    {
        return PropertyDiscoveryStrategy.GetValuesFromProperties(Values, Values.TiedValueNames.ToArray());
    }

    public IEnumerable<IOperationDescription> GetMethods()
    {
        foreach (var name in Values.TiedValueNames)
        {
            if (Values.IsTied(name))
            {
                string valueName = name;
                // in the box of the value's row, right under it (Distance, Direction)
                var group = Values.GetType().GetProperty(name)?.GetAttribute<PropertyGridGroupAttribute>()?.Name;
                yield return new DelegateOperation(
                    TiedValues.UntieVerb(name),
                    "Type the " + name.ToLowerInvariant(),
                    PropertyGridIcon.Pencil,
                    () => Untie(valueName),
                    group);
            }
        }

        yield return MethodDescription.Get<TiedValuesPanel>(nameof(OK));
    }

    // the row becomes editable: the panel is shown again for that
    void Untie(string name)
    {
        if (Values.Detach(name))
        {
            Values.Drawing.RaiseDisplayProperties(this);
        }
    }

    /// <summary>
    /// Back to the tool's panel; the figure stays as it is. Closed first, so that the tool
    /// takes the values as its defaults (<see cref="FigureCreator.ForgetCreatedFigure"/>)
    /// before its own panel reads them.
    /// </summary>
    [PropertyGridIcon(PropertyGridIcon.Check)]
    public void OK()
    {
        var drawing = Values.Drawing;
        drawing.RaiseDisplayProperties(null);
        drawing.RaiseDisplayProperties(drawing.Behavior?.PropertyBag);
        drawing.ClearStatus();
    }

    public override string ToString()
    {
        return Shown.Title;
    }
}
