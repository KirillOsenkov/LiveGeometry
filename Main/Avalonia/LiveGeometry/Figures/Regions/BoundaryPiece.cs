using System;
using System.Collections.Generic;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// A stretch of a region's boundary (<see cref="RegionBoundary"/>): a straight segment or an
/// arc of an ellipse (a circle among them), run through by a parameter from 0 at its start
/// to 1 at its end. Oriented: the region it bounds is on its left.
/// </summary>
public abstract class BoundaryPiece
{
    public abstract Point Start { get; }

    public abstract Point End { get; }

    public abstract Point PointAt(double t);

    /// <summary>The unit direction of travel at the parameter</summary>
    public abstract Point DirectionAt(double t);

    public abstract double Length { get; }

    /// <summary>
    /// The piece's share of the area to its left, by Green's theorem: the integral of
    /// (x - ox) dy - (y - oy) dx over the piece, halved. Taken from an origin near the
    /// figures, so that a tiny piece misjudged costs a tiny area and not one the size of
    /// the coordinates.
    /// </summary>
    public abstract double AreaTerm(Point origin);

    /// <summary>The stretch between the two parameters</summary>
    public abstract BoundaryPiece Part(double t0, double t1);

    public abstract BoundaryPiece Reversed();

    /// <summary>The point of the piece nearest to the point</summary>
    public abstract Point NearestPoint(Point point);

    /// <summary>A box around the piece, for a quick no before the crossings are worked out</summary>
    public abstract Rect Bounds { get; }

    /// <summary>Points along the piece after its start, the end last: a polygon that stands in for it</summary>
    public abstract void AddSamples(List<Point> points);

    public Point Middle => PointAt(0.5);

    /// <summary>The unit normal to the left of the direction of travel at the parameter</summary>
    public Point LeftNormalAt(double t)
    {
        var direction = DirectionAt(t);
        return new Point(-direction.Y, direction.X);
    }

    /// <summary>Whether the two pieces are one and the same stretch, run either way</summary>
    public bool IsSameAs(BoundaryPiece other, double tolerance)
    {
        if (!RegionBoundary.Near(Middle, other.Middle, tolerance))
        {
            return false;
        }

        return RegionBoundary.Near(Start, other.Start, tolerance) && RegionBoundary.Near(End, other.End, tolerance)
            || RegionBoundary.Near(Start, other.End, tolerance) && RegionBoundary.Near(End, other.Start, tolerance);
    }
}

public class SegmentPiece : BoundaryPiece
{
    public SegmentPiece(Point start, Point end)
    {
        Start = start;
        End = end;
    }

    public override Point Start { get; }

    public override Point End { get; }

    public override Point PointAt(double t)
    {
        return new Point(Start.X + (End.X - Start.X) * t, Start.Y + (End.Y - Start.Y) * t);
    }

    public override Point DirectionAt(double t)
    {
        double length = Length;
        return length > 0 ? new Point((End.X - Start.X) / length, (End.Y - Start.Y) / length) : new Point(1, 0);
    }

    public override double Length => Start.Distance(End);

    public override double AreaTerm(Point origin)
    {
        return ((Start.X - origin.X) * (End.Y - origin.Y) - (End.X - origin.X) * (Start.Y - origin.Y)) / 2;
    }

    public override BoundaryPiece Part(double t0, double t1)
    {
        return new SegmentPiece(PointAt(t0), PointAt(t1));
    }

    public override BoundaryPiece Reversed()
    {
        return new SegmentPiece(End, Start);
    }

    public override Point NearestPoint(Point point)
    {
        return Math.GetNearestPointOnSegment(point, new PointPair(Start, End));
    }

    public override Rect Bounds
    {
        get
        {
            double left = System.Math.Min(Start.X, End.X);
            double top = System.Math.Min(Start.Y, End.Y);
            return new Rect(left, top, System.Math.Abs(End.X - Start.X), System.Math.Abs(End.Y - Start.Y));
        }
    }

    public override void AddSamples(List<Point> points)
    {
        points.Add(End);
    }

    /// <summary>Where the point lies along the segment, as a parameter (0 at the start, 1 at the end, beyond either outside)</summary>
    public double ParameterOf(Point point)
    {
        double dx = End.X - Start.X;
        double dy = End.Y - Start.Y;
        double lengthSquared = dx * dx + dy * dy;
        return lengthSquared > 0 ? ((point.X - Start.X) * dx + (point.Y - Start.Y) * dy) / lengthSquared : 0;
    }
}

/// <summary>
/// An arc of an ellipse (x = a cos t, y = b sin t in its own axes, turned by the inclination
/// about the center) from the parametric angle <see cref="From"/> to <see cref="To"/>; it
/// runs clockwise when To is below From. A circle has equal axes.
/// </summary>
public class ArcPiece : BoundaryPiece
{
    public ArcPiece(
        Point center,
        double semiMajor,
        double semiMinor,
        double inclination,
        double from,
        double to)
    {
        // A circle's inclination is only where its parametrization starts (a sector's is
        // the angle to its first point): folded into the angles, so that every arc of one
        // circle has one parametrization. Judged by their inclinations, an arc of a circular
        // segment and an arc of its circle were taken for different ellipses, and the search
        // for crossings between them cut both into slivers at random.
        if (semiMajor == semiMinor && inclination != 0)
        {
            from += inclination;
            to += inclination;
            inclination = 0;
        }

        Center = center;
        SemiMajor = semiMajor;
        SemiMinor = semiMinor;
        Inclination = inclination;
        From = from;
        To = to;
    }

    public Point Center { get; }

    public double SemiMajor { get; }

    public double SemiMinor { get; }

    public double Inclination { get; }

    public double From { get; }

    public double To { get; }

    /// <summary>The angle the arc goes through, positive counterclockwise</summary>
    public double Sweep => To - From;

    public bool IsCounterclockwise => To >= From;

    public bool IsCircle => SemiMajor == SemiMinor;

    public override Point Start => At(From);

    public override Point End => At(To);

    /// <summary>The point of the ellipse at the parametric angle</summary>
    public Point At(double theta)
    {
        return Math.PointOnEllipse(Center, SemiMajor, SemiMinor, Inclination, theta);
    }

    public override Point PointAt(double t)
    {
        return At(From + (To - From) * t);
    }

    public override Point DirectionAt(double t)
    {
        double theta = From + (To - From) * t;
        double cos = System.Math.Cos(Inclination);
        double sin = System.Math.Sin(Inclination);
        double dx = -SemiMajor * System.Math.Sin(theta);
        double dy = SemiMinor * System.Math.Cos(theta);
        double x = dx * cos - dy * sin;
        double y = dx * sin + dy * cos;
        double length = System.Math.Sqrt(x * x + y * y);
        if (!(length > 0))
        {
            return new Point(1, 0);
        }

        double sign = IsCounterclockwise ? 1 : -1;
        return new Point(sign * x / length, sign * y / length);
    }

    public override double Length => Math.EllipseArcLength(SemiMajor, SemiMinor, System.Math.Min(From, To), System.Math.Abs(To - From));

    public override double AreaTerm(Point origin)
    {
        var start = Start;
        var end = End;
        return (SemiMajor * SemiMinor * (To - From)
            + (Center.X - origin.X) * (end.Y - start.Y)
            - (Center.Y - origin.Y) * (end.X - start.X)) / 2;
    }

    public override BoundaryPiece Part(double t0, double t1)
    {
        return new ArcPiece(Center, SemiMajor, SemiMinor, Inclination, From + (To - From) * t0, From + (To - From) * t1);
    }

    public override BoundaryPiece Reversed()
    {
        return new ArcPiece(Center, SemiMajor, SemiMinor, Inclination, To, From);
    }

    /// <summary>The parametric angle of the point of the ellipse on the ray from the center through the point</summary>
    public double AngleOf(Point point)
    {
        double cos = System.Math.Cos(Inclination);
        double sin = System.Math.Sin(Inclination);
        double dx = point.X - Center.X;
        double dy = point.Y - Center.Y;
        double along = dx * cos + dy * sin;
        double across = -dx * sin + dy * cos;
        return System.Math.Atan2(across / SemiMinor, along / SemiMajor);
    }

    /// <summary>
    /// The left side of the ellipse's equation minus 1 at the point: 0 on the curve, below
    /// inside, above outside
    /// </summary>
    public double Implicit(Point point)
    {
        double cos = System.Math.Cos(Inclination);
        double sin = System.Math.Sin(Inclination);
        double dx = point.X - Center.X;
        double dy = point.Y - Center.Y;
        double along = (dx * cos + dy * sin) / SemiMajor;
        double across = (-dx * sin + dy * cos) / SemiMinor;
        return along * along + across * across - 1;
    }

    /// <summary>
    /// Where the parametric angle is along the arc, as a parameter 0 to 1, or NaN when the
    /// arc doesn't reach it (within the angular tolerance past its ends)
    /// </summary>
    public double ParameterOf(double theta, double tolerance)
    {
        double sweep = System.Math.Abs(To - From);
        double u = IsCounterclockwise ? theta - From : From - theta;
        u = u % Math.DOUBLEPI;
        if (u < 0)
        {
            u += Math.DOUBLEPI;
        }

        if (u > Math.DOUBLEPI - tolerance)
        {
            u -= Math.DOUBLEPI;
        }

        if (u < -tolerance || u > sweep + tolerance)
        {
            return double.NaN;
        }

        return sweep > 0 ? System.Math.Clamp(u / sweep, 0, 1) : 0;
    }

    /// <summary>The point of the arc on the ray from the center through the point, or the nearer end when the arc doesn't reach it (exact for a circle)</summary>
    public override Point NearestPoint(Point point)
    {
        double theta = AngleOf(point);
        if (!double.IsNaN(ParameterOf(theta, tolerance: 0)))
        {
            return At(theta);
        }

        var start = Start;
        var end = End;
        return start.Distance(point) <= end.Distance(point) ? start : end;
    }

    /// <summary>The box of the whole ellipse: good enough for a quick no</summary>
    public override Rect Bounds
    {
        get
        {
            double cos = System.Math.Cos(Inclination);
            double sin = System.Math.Sin(Inclination);
            double width = System.Math.Sqrt((SemiMajor * cos).Sqr() + (SemiMinor * sin).Sqr());
            double height = System.Math.Sqrt((SemiMajor * sin).Sqr() + (SemiMinor * cos).Sqr());
            return new Rect(Center.X - width, Center.Y - height, 2 * width, 2 * height);
        }
    }

    public override void AddSamples(List<Point> points)
    {
        int count = System.Math.Max(2, (int)System.Math.Ceiling(System.Math.Abs(To - From) / (Math.PI / 24)));
        for (int i = 1; i <= count; i++)
        {
            points.Add(PointAt((double)i / count));
        }
    }
}
