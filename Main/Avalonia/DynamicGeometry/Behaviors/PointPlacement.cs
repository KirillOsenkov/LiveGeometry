using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace DynamicGeometry;

public enum PointPlacementKind
{
    /// <summary>A new free point</summary>
    Free,

    /// <summary>There already is a point here; nothing new gets created</summary>
    Existing,

    /// <summary>A new point that slides along a line or a circle</summary>
    OnFigure,

    /// <summary>A new point at the intersection of two figures</summary>
    Intersection,

    /// <summary>A new midpoint of a segment</summary>
    Midpoint
}

/// <summary>
/// What a click at some place gives when a tool needs a point there. Both the click and the
/// hover preview (<see cref="ClickPreview"/>) ask <see cref="Find"/>, so the preview can't
/// promise something the click doesn't do.
/// </summary>
public class PointPlacement
{
    static readonly IFigure[] noFigures = new IFigure[0];

    PointPlacement(
        PointPlacementKind kind,
        Point coordinates,
        IReadOnlyList<IFigure> sources)
    {
        Kind = kind;
        Coordinates = coordinates;
        Sources = sources;
    }

    public PointPlacementKind Kind { get; }

    /// <summary>Where the point ends up (logical), which may not be where the cursor is</summary>
    public Point Coordinates { get; }

    /// <summary>The figures the new point is going to depend on</summary>
    public IReadOnlyList<IFigure> Sources { get; }

    public IPoint ExistingPoint { get; private set; }

    /// <summary>True if the new point is tied to the figures under it</summary>
    public bool IsDependent => Kind != PointPlacementKind.Free && Kind != PointPlacementKind.Existing;

    public static PointPlacement Free(Point coordinates)
    {
        return new PointPlacement(PointPlacementKind.Free, coordinates, noFigures);
    }

    public static PointPlacement Existing(IPoint point)
    {
        return new PointPlacement(PointPlacementKind.Existing, point.Coordinates, noFigures)
        {
            ExistingPoint = point
        };
    }

    public static PointPlacement Midpoint(Segment segment)
    {
        return new PointPlacement(PointPlacementKind.Midpoint, segment.Coordinates.Midpoint, new IFigure[] { segment });
    }

    /// <param name="coordinates">Logical coordinates of the click, after snapping</param>
    /// <param name="snapToMidpoint">A click anywhere on a segment means its midpoint</param>
    /// <param name="canUse">Says which figures a point may be put on; null for all visible ones</param>
    public static PointPlacement Find(
        Drawing drawing,
        Point coordinates,
        bool snapToMidpoint,
        Predicate<IFigure> canUse = null)
    {
        var underCursor = drawing.Figures.HitTestMany(coordinates)
            .Where(f => f.IsHitTestVisible && (canUse == null || canUse(f)))
            .Reverse()
            .ToArray();

        var existing = underCursor.OfType<IPoint>().FirstOrDefault();
        if (existing != null)
        {
            return Existing(existing);
        }

        var linear = underCursor.Where(PointOnFigure.CanBeOnFigure).ToArray();
        if (linear.Length == 0)
        {
            return Free(coordinates);
        }

        var intersection = FindIntersection(linear, coordinates, maxDistance: 3 * drawing.CoordinateSystem.CursorTolerance);
        if (intersection != null)
        {
            return intersection;
        }

        // the middle of a segment attracts the point even without snapping;
        // with snapping the whole segment does
        var midpointReach = MidpointReach * drawing.CoordinateSystem.CursorTolerance;
        var segment = linear.OfType<Segment>().FirstOrDefault(s =>
            HasMidpoint(s)
            && (snapToMidpoint || s.Coordinates.Midpoint.Distance(coordinates) <= midpointReach));
        if (segment != null)
        {
            var existingMidpoint = FindExistingMidpoint(segment.Dependencies[0], segment.Dependencies[1]);
            if (existingMidpoint != null)
            {
                return Existing(existingMidpoint);
            }

            return Midpoint(segment);
        }

        var figure = (ILinearFigure)linear[0];
        var onFigure = figure.GetPointFromParameter(figure.GetNearestParameterFromPoint(coordinates));
        if (!onFigure.Exists())
        {
            return Free(coordinates);
        }

        return new PointPlacement(PointPlacementKind.OnFigure, onFigure, new[] { linear[0] });
    }

    /// <summary>In cursor tolerances: how near the middle of a segment counts as "the midpoint"</summary>
    public static double MidpointReach = 2;

    /// <summary>
    /// The (visible) midpoint of these two points, if the drawing already has one
    /// </summary>
    public static MidPoint FindExistingMidpoint(IFigure first, IFigure second)
    {
        return first.Dependents
            .OfType<MidPoint>()
            .FirstOrDefault(m => m.Visible && m.Dependencies.Count == 2 && m.Dependencies.Contains(second));
    }

    public static bool HasMidpoint(Segment segment)
    {
        return segment.Dependencies.Count == 2 && segment.Dependencies.All(d => d is IPoint);
    }

    /// <summary>
    /// Any number of figures can pass under the cursor; the pair that actually crosses nearest
    /// to it wins.
    /// </summary>
    static PointPlacement FindIntersection(IFigure[] figures, Point coordinates, double maxDistance)
    {
        PointPlacement best = null;
        double bestDistance = maxDistance;

        for (int i = 0; i < figures.Length; i++)
        {
            for (int j = i + 1; j < figures.Length; j++)
            {
                var first = figures[i];
                var second = figures[j];
                if (!IntersectionAlgorithms.CanIntersect(first, second))
                {
                    continue;
                }

                var algorithm = IntersectionPoint.DoubleDispatchIntersectionAlgorithm(first, second, coordinates);
                if (algorithm == null)
                {
                    continue;
                }

                var point = algorithm(first, second);
                if (!point.Exists() || first.HitTest(point) == null || second.HitTest(point) == null)
                {
                    continue;
                }

                var distance = point.Distance(coordinates);
                if (distance <= bestDistance)
                {
                    bestDistance = distance;
                    best = new PointPlacement(PointPlacementKind.Intersection, point, new[] { first, second });
                }
            }
        }

        return best;
    }

    /// <summary>
    /// Creates the point (not added to the drawing yet). Null for <see cref="PointPlacementKind.Existing"/>.
    /// </summary>
    public IPoint Create(Drawing drawing)
    {
        switch (Kind)
        {
            case PointPlacementKind.Free:
                return Factory.CreateFreePoint(drawing, Coordinates);
            case PointPlacementKind.OnFigure:
                return Factory.CreatePointOnFigure(drawing, Sources[0], Coordinates);
            case PointPlacementKind.Intersection:
                return Factory.CreateIntersectionPoint(drawing, Sources[0], Sources[1], Coordinates);
            case PointPlacementKind.Midpoint:
                // a list of its own: the segment keeps using (and may change) the one it has
                return Factory.CreateMidPoint(drawing, ((Segment)Sources[0]).Dependencies.ToList());
            default:
                return null;
        }
    }

    /// <summary>True if the other placement would give the same point from the same figures</summary>
    public bool HasSameSources(PointPlacement other)
    {
        return other != null && Kind == other.Kind && Sources.SequenceEqual(other.Sources);
    }
}
