using System;
using System.Collections.Generic;
using System.Linq;
using GuiLabs.Undo;

namespace DynamicGeometry;

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
}

public static class TiedValues
{
    /// <summary>Whether "Type the ..." applies: the value comes from something other than a Number</summary>
    public static bool IsTied(this ITiedValues figure, string name)
    {
        var source = figure.GetSource(name);
        return source != null && !(source is Number);
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
            if (!drawing.Figures.Contains(source))
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
