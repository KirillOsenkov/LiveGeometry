using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

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
        return TiedValues.Rows(Values);
    }

    public IEnumerable<IOperationDescription> GetMethods()
    {
        // the row becomes editable, or a number shows: the panel is shown again for that
        foreach (var button in TiedValues.Buttons(Values, () => Values.Drawing.RaiseDisplayProperties(this)))
        {
            yield return button;
        }

        // a line at an angle: Horizontal and Vertical under its angle, as a segment's panel
        // has them under its direction (the line's own verbs would show its own grid, which
        // ends the panel and the click that ties a value)
        if (Values is LineAtAngle line)
        {
            var group = typeof(LineAtAngle).GetProperty(nameof(LineAtAngle.Angle))?.GetAttribute<PropertyGridGroupAttribute>()?.Name;
            if (line.CanAlign(axis: 0))
            {
                yield return new DelegateOperation(
                    nameof(LineAtAngle.MakeHorizontal),
                    "Horizontal",
                    PropertyGridIcon.Horizontal,
                    () => Align(line, axis: 0),
                    group);
            }

            if (line.CanAlign(axis: 90))
            {
                yield return new DelegateOperation(
                    nameof(LineAtAngle.MakeVertical),
                    "Vertical",
                    PropertyGridIcon.Vertical,
                    () => Align(line, axis: 90),
                    group);
            }
        }

        yield return MethodDescription.Get<TiedValuesPanel>(nameof(OK));
    }

    void Align(LineAtAngle line, double axis)
    {
        line.Align(axis);
        Values.Drawing.RaiseDisplayProperties(this);
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
        drawing.ShowBehaviorHint();
    }

    public override string ToString()
    {
        return Shown.Title;
    }
}
