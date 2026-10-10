using System.Collections.Generic;

namespace DynamicGeometry;

/// <summary>
/// The angle a figure stands for, as the Angle and Angle Bisector tools take it with one
/// click instead of three: the vertex and the two side points, and which of the two angles
/// between them (<see cref="AngleSweep"/>). An angle's mark or number gives its own three
/// points and sweep; an arc, a sector or a segment of a circle or an ellipse gives its
/// center and its begin and end points, so the angle measured is the arc's central angle,
/// the way round the arc goes. The tool builds on the points, as it would on three clicks,
/// and copies the sweep once (the precedent: Circle by Radius takes a segment for its two
/// points).
/// </summary>
public class AnglePoints
{
    public IFigure Vertex { get; private set; }

    public IFigure First { get; private set; }

    public IFigure Second { get; private set; }

    public AngleSweep Sweep { get; private set; }

    /// <summary>The vertex, then the two sides, as the angle's figures depend on them</summary>
    public IList<IFigure> Points => new[] { Vertex, First, Second };

    /// <summary>The angle the figure stands for, or null for a figure that stands for none</summary>
    public static AnglePoints From(IFigure figure)
    {
        var dependencies = figure?.Dependencies;
        switch (figure)
        {
            case AngleArc _:
            case AngleMeasurement _:
                return dependencies.Count == 3
                    ? Create(dependencies[0], dependencies[1], dependencies[2], ((IHasSweep)figure).Sweep)
                    : null;
            case EllipseArcBase arc:
                return dependencies.Count > System.Math.Max(arc.BeginPointIndex, arc.EndPointIndex)
                    ? Create(dependencies[0], dependencies[arc.BeginPointIndex], dependencies[arc.EndPointIndex], arc.Sweep)
                    : null;
            default:
                return null;
        }
    }

    /// <summary>Whether a click on the figure gives an angle, mark and number included</summary>
    public static bool Takes(IFigure figure)
    {
        return From(figure) != null;
    }

    /// <summary>Whether a click on the figure gives an angle that isn't one measured already: an arc-like figure</summary>
    public static bool TakesArc(IFigure figure)
    {
        return !(figure is AngleArc) && !(figure is AngleMeasurement) && From(figure) != null;
    }

    static AnglePoints Create(IFigure vertex, IFigure first, IFigure second, AngleSweep sweep)
    {
        if (!(vertex is IPoint) || !(first is IPoint) || !(second is IPoint) || first == second)
        {
            return null;
        }

        return new AnglePoints()
        {
            Vertex = vertex,
            First = first,
            Second = second,
            Sweep = sweep
        };
    }
}
