using System;
using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry;

/// <summary>
/// Which of the things a click at one place could mean is the one it means. Where figures
/// overlap - a segment along a line, two lines crossing, a polygon's side on its outline -
/// a click takes the first of them as hit testing orders them (the topmost, the nearest
/// point, the newest), and Tab goes on to the next (Shift+Tab back): the hover shows the one
/// a click takes now, and the status says which of how many it is. A finger has no hover,
/// so a tap where there is a choice asks in a menu.
/// <para>
/// A tool says what a click at the cursor could take, best first (<see cref="Behavior.FindClickOptions"/>):
/// figures, or the points it could make (<see cref="PointPlacement"/>). The hover offers
/// them here (<see cref="Offer"/>), and whatever finds what a click takes picks from its own
/// list through <see cref="Pick{T}"/>: the first of it while nothing was chosen, which is
/// what a click took before there was a choice; the chosen one when it is in the list, and
/// nothing when the choice is something another list holds (a point on the segment the
/// Distance tool would otherwise measure). A choice holds for one click and for the place
/// it was made: the next click and anything else under the cursor start over at the first.
/// </para>
/// </summary>
public class ClickChoice
{
    IReadOnlyList<object> offered = Array.Empty<object>();
    int index;

    /// <summary>How many things a click at the cursor could mean</summary>
    public int Count
    {
        get
        {
            return offered.Count;
        }
    }

    /// <summary>Which of them a click takes, from 0</summary>
    public int Index
    {
        get
        {
            return index;
        }
    }

    /// <summary>What a click takes, of what was offered; null when nothing was</summary>
    public object Current
    {
        get
        {
            return offered.Count > 0 ? offered[index] : null;
        }
    }

    /// <summary>
    /// What a click at the cursor could mean, best first. The same as before (the cursor still
    /// over the same figures) keeps the choice; anything else starts at the first.
    /// </summary>
    public void Offer(IEnumerable<object> options)
    {
        var list = options.ToList();
        if (list.Count == offered.Count && list.Zip(offered, Same).All(same => same))
        {
            // the same options, worked out anew: a placement on a figure is somewhere else
            // along it now
            offered = list;
            return;
        }

        offered = list;
        index = 0;
    }

    /// <summary>The next (or with a negative step the one before), round; false when there is nothing to choose from</summary>
    public bool Step(int step)
    {
        if (offered.Count < 2)
        {
            return false;
        }

        index = ((index + step) % offered.Count + offered.Count) % offered.Count;
        return true;
    }

    /// <summary>Takes this option, one of those offered (a tap's menu)</summary>
    public void Choose(object option)
    {
        int found = offered.ToList().FindIndex(item => Same(item, option));
        if (found >= 0)
        {
            index = found;
        }
    }

    /// <summary>Back to the first, nothing offered: a click was made, the cursor left</summary>
    public void Forget()
    {
        offered = Array.Empty<object>();
        index = 0;
    }

    /// <summary>
    /// What a click takes of one list of what it could take (the figures a step wants, the
    /// points it could make): the first while nothing was chosen; else the chosen one, or
    /// null when it is not in this list.
    /// </summary>
    public T Pick<T>(IReadOnlyList<T> options) where T : class
    {
        if (options.Count == 0)
        {
            return null;
        }

        if (index == 0)
        {
            return options[0];
        }

        var current = Current;
        return options.FirstOrDefault(option => Same(option, current));
    }

    /// <summary>
    /// Whether two options are the same thing to click: the same figure; a placement on the
    /// same figures (wherever along them the cursor has it); an existing point and its placement
    /// </summary>
    public static bool Same(object first, object second)
    {
        if (first is PointPlacement { ExistingPoint: not null } existing)
        {
            first = existing.ExistingPoint;
        }

        if (second is PointPlacement { ExistingPoint: not null } other)
        {
            second = other.ExistingPoint;
        }

        if (first is PointPlacement firstPlacement && second is PointPlacement secondPlacement)
        {
            return firstPlacement.HasSameSources(secondPlacement);
        }

        return ReferenceEquals(first, second);
    }

    /// <summary>What an option is, in words: "Segment AB", "A point on line g", "Intersection of circle c and line g"</summary>
    public static string Describe(object option)
    {
        switch (option)
        {
            case PointPlacement { ExistingPoint: not null } placement:
                return Describe(placement.ExistingPoint);
            case PointPlacement placement:
                switch (placement.Kind)
                {
                    case PointPlacementKind.OnFigure:
                        return "A point on " + ConstructionText.Of(placement.Sources[0]);
                    case PointPlacementKind.Intersection:
                        return "Intersection of " + ConstructionText.Of(placement.Sources[0]) + " and " + ConstructionText.Of(placement.Sources[1]);
                    case PointPlacementKind.Midpoint:
                        return "The midpoint of " + ConstructionText.Of(placement.Sources[0]);
                    default:
                        return "A free point";
                }

            case IFigurePart part:
                var words = ConstructionText.Of(part);
                return words.Length > 0 ? char.ToUpperInvariant(words[0]) + words.Substring(1) : words;
            case FigureBase figure:
                return figure.Title;
            case IFigure figure:
                return ConstructionText.Of(figure);
            default:
                return "";
        }
    }

    /// <summary>
    /// An option with how it is built, for the status: "Midpoint E of AB", "Perpendicular
    /// line h to segment AB through C"; <see cref="Describe"/> for anything else
    /// </summary>
    public static string DescribeInFull(object option)
    {
        var words = Describe(option);
        if (option is FigureBase figure && !string.IsNullOrEmpty(figure.Construction))
        {
            words += " " + figure.Construction;
        }

        return words;
    }

    /// <summary>
    /// Whether the status names what a click takes also when there is nothing else to choose
    /// (the Drag tool: whatever is under the cursor); off, it speaks only of a choice
    /// </summary>
    public bool DescribesSingle { get; set; }

    /// <summary>
    /// For the status bar while there is a choice: "Segment AB (2 of 3): Tab for the next.";
    /// the one option alone with <see cref="DescribesSingle"/>; null while there is nothing
    /// </summary>
    public string StatusText
    {
        get
        {
            if (offered.Count == 0 || (offered.Count == 1 && !DescribesSingle))
            {
                return null;
            }

            var words = DescribeInFull(Current);
            if (offered.Count == 1)
            {
                return words;
            }

            return words + " (" + (index + 1) + " of " + offered.Count + "): Tab for the next.";
        }
    }
}
