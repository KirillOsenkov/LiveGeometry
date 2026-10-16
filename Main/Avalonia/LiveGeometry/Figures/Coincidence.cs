using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// Whether a new figure lies on a figure the drawing shows, as it is built: the same circle
/// (one center, and a point of one on the other, or the same radius), the same ellipse (the
/// same three points), the same line (both points of one on the other) - and drawn over a
/// part of it right now. Drawn in the same style, it would be invisible there, so it gets
/// one that shows through (<see cref="StyleManager.CoincidentLineStyleName"/>), decided
/// once when it is made. Only what the construction says counts, never figures that lie on
/// each other by chance: that is what a drag changes.
/// </summary>
public static class Coincidence
{
    /// <summary>How far a proof follows points on figures to figures they lie on, and so on</summary>
    const int MaximumDepth = 3;

    /// <summary>
    /// A figure about to be added gets the style that shows through where it would have the
    /// default line and lies on a figure the drawing shows (see <see cref="Coincidence"/>).
    /// Asked again of the figure a tool previews, as its last point follows the cursor: it
    /// trades that style and the default back and forth, as the figure the click would make.
    /// A style of any other kind is left alone.
    /// </summary>
    public static void StyleIfCoinciding(Drawing drawing, IFigure figure)
    {
        var styles = drawing.StyleManager;
        var defaultStyle = styles.AssignDefaultStyle(figure);
        var coincident = styles.GetStyle(StyleManager.CoincidentLineStyleName);
        if (defaultStyle?.Name != StyleManager.LineStyleName
            || coincident == null
            || !coincident.GetType().SupportsFigureType(figure.GetType())
            || figure.Style != null && figure.Style != defaultStyle && figure.Style != coincident)
        {
            return;
        }

        if (LiesOnShownFigure(drawing, figure))
        {
            if (figure.Style != coincident)
            {
                figure.Style = coincident;
            }
        }
        else if (figure.Style == coincident)
        {
            figure.Style = defaultStyle;
        }
    }

    /// <summary>Whether the figure lies on a figure of the drawing that is shown, as it is built, and over a part of it now</summary>
    public static bool LiesOnShownFigure(Drawing drawing, IFigure figure)
    {
        if (!IsCurve(figure))
        {
            return false;
        }

        var forms = new LinearForms();
        return drawing.Figures.Any(other => other != figure
            && other.Visible
            && other.Exists
            && IsCurve(other)
            && Coincide(figure, other, forms, MaximumDepth)
            && Overlap(figure, other));
    }

    /// <summary>The kinds this knows: lines of every kind, circles, ellipses and their arcs; not an angle's mark, sized in pixels</summary>
    static bool IsCurve(IFigure figure)
    {
        return figure is ILine || figure is IEllipse && figure is not AngleArc;
    }

    static bool Coincide(IFigure first, IFigure second, LinearForms forms, int depth)
    {
        if (first == second)
        {
            return true;
        }

        if (depth == 0)
        {
            return false;
        }

        if (first is ICircle firstCircle && second is ICircle secondCircle)
        {
            return SameCircle(firstCircle, secondCircle, forms, depth);
        }

        if (first is ILine && second is ILine)
        {
            return SameLine(first, second, forms, depth);
        }

        if (first is IEllipse && second is IEllipse && first is not ICircle && second is not ICircle)
        {
            return SameEllipse(first, second, forms);
        }

        return false;
    }

    #region Circles

    static bool SameCircle(ICircle first, ICircle second, LinearForms forms, int depth)
    {
        var firstCenter = CenterOf(first);
        var secondCenter = CenterOf(second);
        if (firstCenter == null || secondCenter == null || !SamePoint(firstCenter, secondCenter, forms))
        {
            return false;
        }

        return ThroughPoints(first, firstCenter, forms).Any(point => LiesOn(point, second, forms, depth - 1))
            || ThroughPoints(second, secondCenter, forms).Any(point => LiesOn(point, first, forms, depth - 1))
            || SameRadius(first, second);
    }

    /// <summary>The center a circle is built on; none for a circle by equation or an angle's mark</summary>
    static IPoint CenterOf(ICircle circle)
    {
        return circle switch
        {
            AngleArc => null,
            CircleByRadius byRadius => byRadius.Dependencies.LastOrDefault() as IPoint,
            Circle or CircleArcBase => circle.Dependencies.FirstOrDefault() as IPoint,
            _ => null
        };
    }

    /// <summary>
    /// The points a circle is built to go through: a circle's or an arc's second point, the
    /// end of a radius from the center (two points, or a segment, one of them the center)
    /// </summary>
    static IEnumerable<IPoint> ThroughPoints(ICircle circle, IPoint center, LinearForms forms)
    {
        var dependencies = circle.Dependencies;
        if (circle is Circle or CircleArcBase)
        {
            if (dependencies.Count > 1 && dependencies[1] is IPoint point)
            {
                yield return point;
            }

            yield break;
        }

        if (circle is not CircleByRadius)
        {
            yield break;
        }

        var radius = dependencies.Count > 2 ? dependencies.Take(2).ToList()
            : dependencies.Count == 2 && dependencies[0] is Segment segment ? segment.Dependencies.ToList()
            : null;
        if (radius == null || radius.Count != 2)
        {
            yield break;
        }

        for (int i = 0; i < 2; i++)
        {
            if (radius[i] is IPoint pivot && radius[1 - i] is IPoint end && SamePoint(pivot, center, forms))
            {
                yield return end;
            }
        }
    }

    /// <summary>Two circles by radius with the radius given by the same figure, or by the same two points</summary>
    static bool SameRadius(ICircle first, ICircle second)
    {
        if (first is not CircleByRadius || second is not CircleByRadius)
        {
            return false;
        }

        var firstRadius = first.Dependencies.Take(first.Dependencies.Count - 1).ToList();
        var secondRadius = second.Dependencies.Take(second.Dependencies.Count - 1).ToList();
        return firstRadius.Count == secondRadius.Count
            && (firstRadius.SequenceEqual(secondRadius) || Enumerable.Reverse(firstRadius).SequenceEqual(secondRadius));
    }

    #endregion

    /// <summary>An ellipse or an elliptical arc on the same center and ends of the axes</summary>
    static bool SameEllipse(IFigure first, IFigure second, LinearForms forms)
    {
        if (first.Dependencies.Count < 3 || second.Dependencies.Count < 3)
        {
            return false;
        }

        for (int i = 0; i < 3; i++)
        {
            if (first.Dependencies[i] is not IPoint firstPoint
                || second.Dependencies[i] is not IPoint secondPoint
                || !SamePoint(firstPoint, secondPoint, forms))
            {
                return false;
            }
        }

        return true;
    }

    #region Lines

    static bool SameLine(IFigure first, IFigure second, LinearForms forms, int depth)
    {
        var points = TwoPoints(first);
        if (points != null)
        {
            return LiesOn(points.Value.First, second, forms, depth - 1)
                && LiesOn(points.Value.Second, second, forms, depth - 1);
        }

        points = TwoPoints(second);
        return points != null
            && LiesOn(points.Value.First, first, forms, depth - 1)
            && LiesOn(points.Value.Second, first, forms, depth - 1);
    }

    /// <summary>The two points a line, ray, segment or vector is drawn through; not a parallel, a perpendicular or a bisector, which are built on other things</summary>
    static (IPoint First, IPoint Second)? TwoPoints(IFigure line)
    {
        bool throughTwoPoints = line switch
        {
            AngleBisector => false,
            PerpendicularLineBase => false,
            ParallelLine => false,
            Segment or Ray or Vector or LineTwoPoints => true,
            _ => false
        };
        if (throughTwoPoints
            && line.Dependencies.Count == 2
            && line.Dependencies[0] is IPoint first
            && line.Dependencies[1] is IPoint second)
        {
            return (first, second);
        }

        return null;
    }

    /// <summary>The point a parallel or a perpendicular is drawn through</summary>
    static IPoint ThroughPoint(IFigure line)
    {
        return line is ParallelLine or PerpendicularLine && line.Dependencies.Count == 2
            ? line.Dependencies[1] as IPoint
            : null;
    }

    #endregion

    /// <summary>
    /// Whether the point is on the figure as it is built: one of the points the figure is
    /// drawn through, a point put on it (or on a figure that coincides with it), a crossing
    /// of it, the midpoint of two points on a line
    /// </summary>
    static bool LiesOn(IPoint point, IFigure figure, LinearForms forms, int depth)
    {
        if (figure is ICircle circle)
        {
            var center = CenterOf(circle);
            if (center != null && ThroughPoints(circle, center, forms).Any(through => SamePoint(point, through, forms)))
            {
                return true;
            }
        }

        if (figure is ILine)
        {
            var points = TwoPoints(figure);
            if (points != null && (SamePoint(point, points.Value.First, forms) || SamePoint(point, points.Value.Second, forms)))
            {
                return true;
            }

            var through = ThroughPoint(figure);
            if (through != null && SamePoint(point, through, forms))
            {
                return true;
            }

            if (point is MidPoint midpoint
                && depth > 0
                && midpoint.Dependencies.Count == 2
                && midpoint.Dependencies.All(end => end is IPoint endPoint && LiesOn(endPoint, figure, forms, depth - 1)))
            {
                return true;
            }
        }

        if (point is PointOnFigure || point is IntersectionPoint)
        {
            return point.Dependencies.Any(carrier => carrier == figure || Coincide(carrier, figure, forms, depth));
        }

        return false;
    }

    /// <summary>The same point, or the same point however the figures move (<see cref="LinearForms"/>)</summary>
    static bool SamePoint(IPoint first, IPoint second, LinearForms forms)
    {
        if (first == second)
        {
            return true;
        }

        var a = first.Coordinates;
        var b = second.Coordinates;
        double tolerance = Math.TangencyTolerance(System.Math.Max(
            System.Math.Max(System.Math.Abs(a.X), System.Math.Abs(a.Y)),
            System.Math.Max(System.Math.Abs(b.X), System.Math.Abs(b.Y))));
        return System.Math.Abs(a.X - b.X) <= tolerance
            && System.Math.Abs(a.Y - b.Y) <= tolerance
            && forms.AreSamePoint(first, second);
    }

    #region Whether they overlap now

    /// <summary>Whether two figures on one circle or line, as proven, are drawn over a common part of it now: arcs and segments may lie apart</summary>
    static bool Overlap(IFigure first, IFigure second)
    {
        if (first is IEllipse && second is IEllipse)
        {
            var (firstStart, firstSweep) = AngularExtent(first);
            var (secondStart, secondSweep) = AngularExtent(second);
            const double epsilon = 1e-9;
            const double fullTurn = 2 * System.Math.PI;
            if (!(firstSweep > epsilon) || !(secondSweep > epsilon))
            {
                return false;
            }

            double offset = System.Math.IEEERemainder(secondStart - firstStart, fullTurn);
            if (offset < 0)
            {
                offset += fullTurn;
            }

            return offset < firstSweep - epsilon || offset + secondSweep > fullTurn + epsilon;
        }

        if (first is ILine firstLine && second is ILine secondLine)
        {
            var origin = firstLine.Coordinates.P1;
            var along = firstLine.Coordinates.P2.Minus(origin);
            double length = System.Math.Sqrt(along.X * along.X + along.Y * along.Y);
            if (!(length > 0))
            {
                return false;
            }

            var direction = new Point(along.X / length, along.Y / length);
            var (firstLow, firstHigh) = Extent(first, firstLine.Coordinates, origin, direction);
            var (secondLow, secondHigh) = Extent(second, secondLine.Coordinates, origin, direction);
            double common = System.Math.Min(firstHigh, secondHigh) - System.Math.Max(firstLow, secondLow);
            return common > Math.TangencyTolerance(length);
        }

        return false;
    }

    /// <summary>Where an arc starts going counterclockwise, and how far it goes; a whole turn for a circle or an ellipse</summary>
    static (double Start, double Sweep) AngularExtent(IFigure figure)
    {
        if (figure is EllipseArcBase arc)
        {
            return (arc.IsClockwise ? arc.EndAngle : arc.StartAngle, arc.Angle);
        }

        return (0, 2 * System.Math.PI);
    }

    /// <summary>The stretch of the line the figure covers, by the distance along the direction from the origin</summary>
    static (double Low, double High) Extent(IFigure figure, PointPair coordinates, Point origin, Point direction)
    {
        double start = Along(coordinates.P1, origin, direction);
        double end = Along(coordinates.P2, origin, direction);
        if (figure is Segment || figure is Vector)
        {
            return (System.Math.Min(start, end), System.Math.Max(start, end));
        }

        if (figure is Ray && !(figure is AngleBisector bisector && bisector.IsLine))
        {
            return end >= start ? (start, double.PositiveInfinity) : (double.NegativeInfinity, start);
        }

        return (double.NegativeInfinity, double.PositiveInfinity);
    }

    static double Along(Point point, Point origin, Point direction)
    {
        return (point.X - origin.X) * direction.X + (point.Y - origin.Y) * direction.Y;
    }

    #endregion
}
