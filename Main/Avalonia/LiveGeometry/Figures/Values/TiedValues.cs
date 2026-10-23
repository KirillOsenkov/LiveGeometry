using System;
using System.Collections.Generic;
using System.Linq;
using GuiLabs.Undo;

namespace LiveGeometry;

/// <summary>
/// A figure some of whose numbers - an angle, a distance, a factor - are each either
/// <em>typed</em>, held by an auxiliary <see cref="Number"/> the figure depends on, or
/// <em>tied</em> to another figure it depends on (an angle measurement, a slider, a segment)
/// and follows. Each is a row of the property grid, named by <see cref="TiedValueNames"/>:
/// editable while typed, captioned with the source's name while tied, with a "Type the ..."
/// button named <c>Untie</c> + name for the way back - all decided by the figure's
/// <see cref="IConditionalProperties"/>. Tying happens with a click on a suitable figure
/// while the panel that follows the construction is up (<see cref="TiedValuesPanel"/>,
/// <see cref="FigureCreator.ShowCreatedFigure"/>).
/// </summary>
public interface ITiedValues : IFigure
{
    IEnumerable<string> TiedValueNames { get; }

    /// <summary>The figure the value comes from: a Number when typed; null when free (a translated point being dragged)</summary>
    IFigure GetSource(string name);

    /// <summary>Whether a click on this figure could tie the value to it</summary>
    bool Accepts(string name, IFigure figure);

    /// <summary>The value follows <paramref name="source"/> from now on, in one undo step; false when refused</summary>
    bool TieTo(string name, IFigure source);

    /// <summary>The way back: a typed value again, at the value it has now; false when it is typed already</summary>
    bool Detach(string name);

    /// <summary>
    /// What the value is to this figure, as the kind of the Number that holds it: "Rotation
    /// angle", "Translation distance", "Dilation factor", "Fixed length" (the Figure List
    /// said "Number n1" of every one of them)
    /// </summary>
    string DescribeValue(string name);
}

/// <summary>
/// What a <see cref="Number"/> is to the figures that take it (<see cref="TiedValues.RoleOf"/>):
/// its kind in words, whether it is an angle (said in degrees), and whose it is, the words
/// after its value in its construction (" of ABC about O", " of line g")
/// </summary>
public record TiedValueRole(string Kind, bool IsAngle, string Subject);

public static class TiedValues
{
    /// <summary>Whether "Type the ..." applies: the value comes from something other than a Number</summary>
    public static bool IsTied(this ITiedValues figure, string name)
    {
        var source = figure.GetSource(name);
        return source != null && !(source is Number);
    }

    /// <summary>Whether the value named is an angle, in degrees: the angle of a rotation or a line, the direction of a translation</summary>
    public static bool IsAngleValue(string name)
    {
        return name == "Angle" || name == "Direction";
    }

    /// <summary>
    /// What the number is to the figures built on it, when every one of them takes it as the
    /// same value (the vertices of one rotated triangle share their angle): null for a
    /// number nothing takes so, or one an expression reads as well
    /// </summary>
    public static TiedValueRole RoleOf(Number number)
    {
        string kind = null;
        bool isAngle = false;
        var owners = new List<ITiedValues>();
        foreach (var dependent in number.Dependents.Distinct())
        {
            if (dependent is not ITiedValues values)
            {
                return null;
            }

            var name = values.TiedValueNames.FirstOrDefault(candidate => values.GetSource(candidate) == number);
            var described = name == null ? null : values.DescribeValue(name);
            if (described == null || (kind != null && described != kind))
            {
                return null;
            }

            kind = described;
            isAngle = IsAngleValue(name);
            owners.Add(values);
        }

        return kind == null ? null : new TiedValueRole(kind, isAngle, Subject(owners));
    }

    /// <summary>
    /// Whose the value is: the figure transformed with its center (" of ABC about O"), the
    /// line at the angle, the segment whose length or direction is fixed
    /// </summary>
    static string Subject(List<ITiedValues> owners)
    {
        switch (owners[0])
        {
            case LineAtAngle line:
                return " of " + ConstructionText.Of(line);
            case BezierPath path:
                return " of " + ConstructionText.Of(path);
            case TranslatedPoint end when end.IsDistanceFree || end.IsDirectionFree:
                // a fixed length or direction: the segment from the pivot to the end, if drawn
                var pivot = end.Source;
                var drawn = pivot == null ? null : end.Drawing?.Figures.FindLine(pivot, end);
                return drawn != null
                    ? " of " + ConstructionText.Of(drawn)
                    : " of " + ConstructionText.Of(end) + " from " + ConstructionText.Of(pivot);
            default:
                return FigureTransformation.OfPoints(owners.Cast<IFigure>().ToList())?.DescribeSubject() ?? "";
        }
    }

    /// <summary>
    /// The rows of the figure's tied values, as a panel or another figure's page shows them
    /// (<see cref="TiedValuesPanel"/>, the page of a figure that is the image of a
    /// transformation): the figure's own, so an edit is an undo step on the figure
    /// </summary>
    public static IEnumerable<IValueProvider> Rows(ITiedValues values)
    {
        return PropertyDiscoveryStrategy.GetValuesFromProperties(values, values.TiedValueNames.ToArray());
    }

    /// <summary>
    /// The buttons under those rows, each in the box of its value's row: "Type the ..." for
    /// a value taken from a figure (<see cref="ITiedValues.Detach"/>), "Show the ... on the
    /// drawing" for one held by a hidden Number (<see cref="ShowNumber"/>)
    /// </summary>
    /// <param name="refresh">Shows the page again after a button changed what the rows are</param>
    public static IEnumerable<IOperationDescription> Buttons(ITiedValues values, Action refresh)
    {
        foreach (var name in values.TiedValueNames)
        {
            string valueName = name;
            // in the box of the value's row, right under it (Distance, Direction)
            var group = values.GetType().GetProperty(name)?.GetAttribute<PropertyGridGroupAttribute>()?.Name;
            if (values.IsTied(name))
            {
                yield return new DelegateOperation(
                    UntieVerb(name),
                    "Type the " + name.ToLowerInvariant(),
                    PropertyGridIcon.Pencil,
                    () =>
                    {
                        if (values.Detach(valueName))
                        {
                            refresh();
                        }
                    },
                    group);
            }
            else if (values.GetSource(name) is Number number && !number.Visible)
            {
                yield return new DelegateOperation(
                    "Show" + name,
                    "Show the " + name.ToLowerInvariant() + " on the drawing",
                    PropertyGridIcon.Eye,
                    () =>
                    {
                        ShowNumber(number);
                        refresh();
                    },
                    group);
            }
        }
    }

    /// <summary>
    /// The number's caption ("a = 45") onto the paper, one undo step: a helper made for its
    /// figure becomes a figure of its own first, named as the Number tool names one (a, b,
    /// c), so that it stays when its figure goes and can be read in a formula by a name that
    /// says it is the user's
    /// </summary>
    public static void ShowNumber(Number number)
    {
        var drawing = number.Drawing;
        var actionManager = drawing.ActionManager;
        using (Transaction.Create(actionManager, delayed: false))
        {
            if (number.Auxiliary)
            {
                Actions.SetProperty(actionManager, new PropertyValue(nameof(FigureBase.Auxiliary), number), false);
                Actions.SetProperty(actionManager, new PropertyValue(nameof(IFigure.Name), number), FigureBase.GenerateLetterName(number, "a", null));
            }

            Actions.SetProperty(actionManager, new PropertyValue(nameof(IFigure.Visible), number), true);
        }
    }

    /// <summary>
    /// What a caption calls the figure a value comes from: its name, or for a part, which
    /// has none (a side of a regular polygon), the name of the figure it is a part of
    /// </summary>
    public static string SourceName(IFigure source)
    {
        if (!string.IsNullOrEmpty(source.Name))
        {
            return source.Name;
        }

        return source.Drawing?.Figures.FindTopLevel(source)?.Name ?? "";
    }

    /// <summary>The name of the "Type the ..." button of a value: UntieAngle</summary>
    public static string UntieVerb(string name)
    {
        return "Untie" + name;
    }

    /// <summary>
    /// Every owner's value (the vertices of one rotated triangle share their angle) from
    /// <paramref name="oldSource"/> to <paramref name="source"/>, in one undo step: the source
    /// is added when it is new (a Number for a typed value), moved before the owners in the
    /// list with what it is built on so that the list stays in dependency order, the
    /// dependency swapped on each owner, and the old source removed when it was a Number made
    /// for them that nothing uses any more. Refused for a source built on an owner, which
    /// would be a cycle.
    /// </summary>
    public static bool Tie(IList<IFigure> owners, IFigure oldSource, IFigure source)
    {
        return Tie(owners, oldSource, source, owner => Actions.ReplaceDependency(owner, oldSource, source));
    }

    /// <param name="retie">Records the swap on one owner, for a figure whose dependency list needs more than a one-for-one replacement</param>
    public static bool Tie(
        IList<IFigure> owners,
        IFigure oldSource,
        IFigure source,
        Action<IFigure> retie)
    {
        if (owners.Count == 0 || source == null || source == oldSource || owners.Any(owner => source.DependsOn(owner)))
        {
            return false;
        }

        var drawing = owners[0].Drawing;
        using (Transaction.Create(drawing.ActionManager, false))
        {
            // (a side of a regular polygon is in the drawing as a part of the polygon)
            if (drawing.Figures.FindTopLevel(source) == null)
            {
                Actions.Add(drawing, source);
            }

            foreach (var owner in owners)
            {
                Actions.MoveBefore(drawing, source, owner);
            }

            foreach (var owner in owners)
            {
                retie(owner);
            }

            if (oldSource is Number number && number.Auxiliary && number.Dependents.IsEmpty())
            {
                Actions.Remove(number);
            }
        }

        return true;
    }
}
