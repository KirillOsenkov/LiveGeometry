using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// An angle that a click near its vertex measures in one go: the Angle tool takes the vertex
/// and a point on each side without three clicks when two lines (segments, rays, sides of a
/// polygon) leave a point and the cursor is inside the mark the angle would get. Only when
/// that is the one angle the cursor can mean: with a third line through the vertex on the
/// same side (a bisector), the small angle and the large one both have the cursor inside,
/// and nothing is offered.
/// </summary>
public class AngleAtVertex
{
    public IPoint Vertex { get; private set; }

    /// <summary>On the side the angle goes counterclockwise from; under 180 degrees to <see cref="Second"/></summary>
    public IPoint First { get; private set; }

    public IPoint Second { get; private set; }

    /// <summary>What the two sides are drawn as (for the halos of the preview)</summary>
    public IReadOnlyList<IFigure> SideFigures { get; private set; }

    /// <summary>The vertex, then the two sides, as the angle's figures depend on them</summary>
    public IList<IFigure> Points => new IFigure[] { Vertex, First, Second };

    public bool IsSameAngle(AngleAtVertex other)
    {
        return other != null && other.Vertex == Vertex && other.First == First && other.Second == Second;
    }

    /// <summary>A direction out of the vertex along something drawn, and a point on it if there is one</summary>
    class Side
    {
        public double Angle;
        public IPoint Point;
        public IFigure Figure;
    }

    /// <summary>
    /// The angle a click at the cursor would measure, or null: none, or more than one (two
    /// vertices in reach, overlapping angles at one vertex, an angle along a line with no
    /// point on it), or the angle there is measured already. Never when a point is under the
    /// cursor: that click is the vertex of an angle measured in three clicks.
    /// </summary>
    public static AngleAtVertex Find(Drawing drawing, Point cursor)
    {
        if (drawing == null || drawing.Figures.HitTest<IPoint>(cursor) != null)
        {
            return null;
        }

        var coordinateSystem = drawing.CoordinateSystem;
        var cursorPixels = coordinateSystem.ToPhysical(cursor);
        var reach = AngleArc.DefaultSize + Math.CursorTolerance;
        AngleAtVertex found = null;
        foreach (var vertex in drawing.Figures.OfType<IPoint>())
        {
            if (!vertex.Visible
                || !vertex.Exists
                || coordinateSystem.ToPhysical(vertex.Coordinates).Distance(cursorPixels) > reach)
            {
                continue;
            }

            var sides = FindSides(drawing, vertex);
            var cursorAngle = Math.GetAngle(vertex.Coordinates, cursor);
            foreach (var first in sides)
            {
                foreach (var second in sides)
                {
                    var sweep = Turn(first.Angle, second.Angle);
                    if (sweep < AngleTolerance
                        || sweep > Math.PI - AngleTolerance
                        || Turn(first.Angle, cursorAngle) > sweep)
                    {
                        continue;
                    }

                    if (found != null || first.Point == null || second.Point == null)
                    {
                        return null;
                    }

                    found = new AngleAtVertex()
                    {
                        Vertex = vertex,
                        First = first.Point,
                        Second = second.Point,
                        SideFigures = new[] { first.Figure, second.Figure }
                    };
                }
            }
        }

        return found == null || IsMeasured(found) ? null : found;
    }

    // Directions that differ by less are one: computed from two points along one line, they
    // come out a hair apart.
    const double AngleTolerance = 1e-9;

    // a fraction of the line's two points' distance
    const double ContinuationStep = 1e-6;

    /// <summary>Counterclockwise from one direction to the other, in [0, 2pi)</summary>
    static double Turn(double from, double to)
    {
        var result = to - from;
        while (result < 0)
        {
            result += 2 * Math.PI;
        }

        while (result >= 2 * Math.PI)
        {
            result -= 2 * Math.PI;
        }

        return 2 * Math.PI - result < AngleTolerance ? 0 : result;
    }

    /// <summary>
    /// The directions in which visible lines, segments, rays, vectors and polygon sides leave
    /// the vertex, one per direction, counterclockwise. A direction with no point on it (a
    /// line through the vertex that runs through no other point) is there too: it parts the
    /// angles around the vertex, but no angle can be measured along it.
    /// </summary>
    static List<Side> FindSides(Drawing drawing, IPoint vertex)
    {
        var sides = new List<Side>();
        foreach (var figure in drawing.Figures)
        {
            if (!figure.Visible || !figure.Exists || figure == vertex)
            {
                continue;
            }

            if (figure is ILine line)
            {
                AddSides(sides, line, vertex);
            }
            else if (figure is Polygon polygon)
            {
                AddSides(sides, polygon, vertex);
            }
        }

        // one side per direction: the first one found, with a point if any has one
        var result = new List<Side>();
        foreach (var side in sides.OrderBy(side => side.Angle))
        {
            var same = result.FirstOrDefault(other => System.Math.Min(Turn(other.Angle, side.Angle), Turn(side.Angle, other.Angle)) < AngleTolerance);
            if (same == null)
            {
                result.Add(side);
            }
            else if (same.Point == null)
            {
                same.Point = side.Point;
                same.Figure = side.Figure;
            }
        }

        return result;
    }

    static void AddSides(List<Side> sides, ILine line, IPoint vertex)
    {
        var at = vertex.Coordinates;
        var coordinates = line.Coordinates;
        var along = coordinates.P2 - coordinates.P1;
        var length = coordinates.P1.Distance(coordinates.P2);
        if (!(length > 0) || !IsOn(line, at))
        {
            return;
        }

        var parameter = line.GetNearestParameterFromPoint(at);
        var parameterTolerance = Tolerance(at, coordinates.P1, coordinates.P2) / length;
        var points = line.Dependencies.OfType<IPoint>()
            .Concat(line.Dependents.OfType<IPoint>())
            .Where(point => point != vertex && point.Exists && IsOn(line, point.Coordinates))
            .ToList();
        foreach (var forward in new[] { true, false })
        {
            // The line goes on from the vertex towards P2 (forward), towards P1, or both: a
            // step past the vertex is still on it, where a segment or a ray that ends there
            // takes the step back to its end. (Not by the parameter domain, which for a line
            // is what is on screen, either way round.)
            var step = forward ? ContinuationStep : -ContinuationStep;
            var stepped = line.GetNearestParameterFromPoint(line.GetPointFromParameter(parameter + step));
            if (System.Math.Abs(stepped - parameter) < ContinuationStep / 2)
            {
                continue;
            }

            var point = points.FirstOrDefault(p =>
            {
                var offset = line.GetNearestParameterFromPoint(p.Coordinates) - parameter;
                return forward ? offset > parameterTolerance : offset < -parameterTolerance;
            });
            sides.Add(new Side()
            {
                Angle = Math.GetAngle(at, forward ? at + along : at - along),
                Point = point,
                Figure = line
            });
        }
    }

    static void AddSides(List<Side> sides, Polygon polygon, IPoint vertex)
    {
        var vertices = polygon.Dependencies.OfType<IPoint>().ToList();
        var index = vertices.IndexOf(vertex);
        if (index < 0 || vertices.Count < 3)
        {
            return;
        }

        var at = vertex.Coordinates;
        foreach (var neighbor in new[] { vertices[(index + vertices.Count - 1) % vertices.Count], vertices[(index + 1) % vertices.Count] })
        {
            if (neighbor.Exists && neighbor.Coordinates.Distance(at) > Tolerance(at, neighbor.Coordinates))
            {
                sides.Add(new Side() { Angle = Math.GetAngle(at, neighbor.Coordinates), Point = neighbor, Figure = polygon });
            }
        }
    }

    /// <summary>Whether a place is on the line, segment or ray itself, but for rounding</summary>
    static bool IsOn(ILine line, Point point)
    {
        var nearest = line.GetPointFromParameter(line.GetNearestParameterFromPoint(point));
        return nearest.Distance(point) <= Tolerance(point, line.Coordinates.P1, line.Coordinates.P2);
    }

    static double Tolerance(params Point[] points)
    {
        var scale = points.Max(point => System.Math.Max(System.Math.Abs(point.X), System.Math.Abs(point.Y)));
        return Math.TangencyTolerance(scale);
    }

    /// <summary>An angle mark or number at the vertex that goes the same way round the same two directions</summary>
    static bool IsMeasured(AngleAtVertex angle)
    {
        var at = angle.Vertex.Coordinates;
        var first = Math.GetAngle(at, angle.First.Coordinates);
        var second = Math.GetAngle(at, angle.Second.Coordinates);
        return angle.Vertex.Dependents.Any(figure =>
            (figure is AngleArc || figure is AngleMeasurementBase)
            && figure.Dependencies.Count == 3
            && figure.Dependencies[0] == angle.Vertex
            && IsSameDirection(Math.GetAngle(at, figure.Point(1)), first)
            && IsSameDirection(Math.GetAngle(at, figure.Point(2)), second));
    }

    static bool IsSameDirection(double angle1, double angle2)
    {
        return System.Math.Min(Turn(angle1, angle2), Turn(angle2, angle1)) < AngleTolerance;
    }
}
