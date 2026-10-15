using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>
/// A figure with parts that other figures can be built on, though the parts are not figures
/// of the drawing themselves: the sides, handles and inside of a Bezier path. A part has no
/// name; a file refers to it by its owner's name and the part's
/// (<c>&lt;Dependency Name="p" Part="In3" /&gt;</c>).
/// </summary>
public interface IFigureParts : IFigure
{
    /// <summary>What a file calls the part; null if the figure is not a part of this one</summary>
    string GetPartName(IFigure part);

    /// <summary>The part a file calls so; null if there is none</summary>
    IFigure GetPart(string partName);

    /// <summary>
    /// The parts a click selects by themselves, to be styled (a path's sides); a click on
    /// any other part (the inside) selects the figure
    /// </summary>
    IEnumerable<IFigure> SelectableParts { get; }
}

/// <summary>A part of a figure (<see cref="IFigureParts"/>)</summary>
public interface IFigurePart : IFigure
{
    IFigureParts Owner { get; }
}

/// <summary>
/// What a selection means where a figure has parts. A part selected by itself is styled by
/// itself; everything else done to the selection (delete, hide, lock, copy) is done to the
/// figure it belongs to.
/// </summary>
public static class FigureParts
{
    /// <summary>What a click on the figure selects: a side by itself, but the inside is the figure</summary>
    public static IFigure SelectionTarget(IFigure figure)
    {
        return figure is IFigurePart part && !part.Owner.SelectableParts.Contains(figure)
            ? part.Owner
            : figure;
    }

    /// <summary>The figure of the drawing that the figure is, or is a part of</summary>
    public static IFigure Whole(IFigure figure)
    {
        return figure is IFigurePart part ? part.Owner : figure;
    }

    /// <summary>The figures of the drawing that these are, or are parts of, each once</summary>
    public static IEnumerable<IFigure> Wholes(IEnumerable<IFigure> figures)
    {
        return figures.Select(Whole).Distinct();
    }

    /// <summary>The figure and its parts that can be selected by themselves</summary>
    public static IEnumerable<IFigure> WithParts(IFigure figure)
    {
        yield return figure;
        if (figure is IFigureParts parts)
        {
            foreach (var part in parts.SelectableParts)
            {
                yield return part;
            }
        }
    }

    /// <summary>Whether the figure is selected, or a part of it by itself</summary>
    public static bool HasSelection(IFigure figure)
    {
        return figure.Selected
            || figure is IFigureParts parts && parts.SelectableParts.Any(part => part.Selected);
    }
}
