using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Media;

namespace LiveGeometry;

/// <summary>What a <see cref="Region"/> makes of its operands</summary>
public enum RegionOperation
{
    /// <summary>The first operand without the others</summary>
    Difference,

    /// <summary>What all the operands share</summary>
    Intersection,

    /// <summary>All the operands together</summary>
    Union
}

/// <summary>
/// The boundary of a set made of other regions by a <see cref="RegionOperation"/>, traced
/// from the operands' boundaries: every operand's boundary is cut into pieces at the
/// crossings with the others (<see cref="BoundaryPiece"/>), and a piece is kept when the
/// result lies on one side of it and not the other - judged at the piece's middle, which
/// is on its own operand's boundary and, the pieces being cut at every crossing, on no
/// other's (a stretch two operands share is found as such). From the kept pieces, oriented
/// with the result on their left, come the area (Green's theorem, exact for segments and
/// arcs), the perimeter, the loops to draw, the centroid and the nearest point of the edge.
/// Nothing is flattened but a Bezier path, whose inside is already the polygon through its
/// samples (<see cref="BezierPath.Area"/>).
/// </summary>
public class RegionBoundary
{
    public static readonly RegionBoundary Empty = new RegionBoundary();

    RegionBoundary()
    {
        Pieces = Array.Empty<BoundaryPiece>();
        Loops = Array.Empty<IReadOnlyList<BoundaryPiece>>();
        Centroid = Math.InfinitePoint;
    }

    /// <summary>The pieces of the boundary, the result on the left of each</summary>
    public IReadOnlyList<BoundaryPiece> Pieces { get; private set; }

    /// <summary>The pieces chained into closed loops, for drawing (even-odd: the loops never overlap, a hole is a loop inside another)</summary>
    public IReadOnlyList<IReadOnlyList<BoundaryPiece>> Loops { get; private set; }

    public double Area { get; private set; }

    public double Perimeter { get; private set; }

    /// <summary>The middle of the set by area; nowhere for an empty one</summary>
    public Point Centroid { get; private set; }

    /// <summary>How big the coordinates are: what the tolerances go by</summary>
    double scale = 1;

    /// <summary>Whether there is an area to speak of</summary>
    public bool HasInside => Pieces.Count > 0 && Area > 1e-12 * scale * scale;

    /// <summary>The point of the boundary nearest to the point; nowhere without a boundary</summary>
    public Point NearestPoint(Point point)
    {
        var result = Math.InfinitePoint;
        double best = double.PositiveInfinity;
        foreach (var piece in Pieces)
        {
            var candidate = piece.NearestPoint(point);
            double distance = candidate.Distance(point);
            if (candidate.Exists() && distance < best)
            {
                best = distance;
                result = candidate;
            }
        }

        return result;
    }

    /// <summary>How far the point is from the boundary; infinite without one</summary>
    public double DistanceTo(Point point)
    {
        var nearest = NearestPoint(point);
        return nearest.Exists() ? nearest.Distance(point) : double.PositiveInfinity;
    }

    /// <summary>Whether the operation holds for the memberships, one per operand in order</summary>
    public static bool Evaluate(RegionOperation operation, IReadOnlyList<bool> memberships)
    {
        if (memberships.Count == 0)
        {
            return false;
        }

        switch (operation)
        {
            case RegionOperation.Difference:
                if (!memberships[0])
                {
                    return false;
                }

                for (int i = 1; i < memberships.Count; i++)
                {
                    if (memberships[i])
                    {
                        return false;
                    }
                }

                return true;
            case RegionOperation.Intersection:
                return memberships.All(m => m);
            default:
                return memberships.Any(m => m);
        }
    }

    /// <summary>Whether the point is in the set the operation makes of the operands</summary>
    public static bool Contains(RegionOperation operation, IReadOnlyList<IRegion> operands, Point point)
    {
        var memberships = new bool[operands.Count];
        for (int i = 0; i < operands.Count; i++)
        {
            var operand = operands[i];
            memberships[i] = operand != null && operand.Exists && operand.HasInside && operand.ContainsPoint(point);
        }

        return Evaluate(operation, memberships);
    }

    public static bool Near(Point a, Point b, double tolerance)
    {
        return System.Math.Abs(a.X - b.X) <= tolerance && System.Math.Abs(a.Y - b.Y) <= tolerance;
    }

    #region Tracing

    /// <summary>The boundary of the set the operation makes of the operands as they are now</summary>
    public static RegionBoundary Trace(RegionOperation operation, IReadOnlyList<IRegion> operands)
    {
        var result = new RegionBoundary();
        if (operands.Count == 0 || operands.Any(operand => operand == null || !operand.Exists))
        {
            return result;
        }

        // every operand's boundary, each piece remembering whose it is
        var pieces = new List<(BoundaryPiece Piece, int Operand)>();
        for (int i = 0; i < operands.Count; i++)
        {
            if (operands[i].HasInside)
            {
                foreach (var piece in PiecesOf(operands[i]))
                {
                    pieces.Add((piece, i));
                }
            }
        }

        if (pieces.Count == 0)
        {
            return result;
        }

        double scale = ScaleOf(pieces.Select(p => p.Piece));
        double tolerance = Tolerance(scale);
        result.scale = scale;

        // cut at every crossing, with the others and with its own operand's (a self-crossing polygon)
        pieces = Split(pieces, tolerance);

        // each piece turned so that its own operand is on its left; one with its operand
        // on neither side or both (two sides of a polygon on top of each other) goes
        var oriented = new List<(BoundaryPiece Piece, int Operand)>();
        foreach (var (piece, operand) in pieces)
        {
            var turned = OrientToOperand(piece, operands[operand], scale);
            if (turned != null)
            {
                oriented.Add((turned, operand));
            }
        }

        // what is on the result's boundary, the result on the left
        var kept = new List<BoundaryPiece>();
        var memberships = new bool[operands.Count];
        foreach (var (piece, operand) in oriented)
        {
            var middle = piece.Middle;
            bool inLeft;
            bool inRight;

            // the other operands as they are at the middle: a stretch shared with another
            // operand goes by that operand's side of it
            for (int j = 0; j < operands.Count; j++)
            {
                memberships[j] = j != operand && IsInside(operands[j], middle);
            }

            foreach (var (other, otherOperand) in oriented)
            {
                if (otherOperand != operand && other.IsSameAs(piece, tolerance))
                {
                    memberships[otherOperand] = Near(other.Start, piece.Start, tolerance);
                }
            }

            memberships[operand] = true;
            inLeft = Evaluate(operation, memberships);
            foreach (var (other, otherOperand) in oriented)
            {
                if (otherOperand != operand && other.IsSameAs(piece, tolerance))
                {
                    memberships[otherOperand] = !Near(other.Start, piece.Start, tolerance);
                }
            }

            memberships[operand] = false;
            inRight = Evaluate(operation, memberships);
            if (inLeft == inRight)
            {
                continue;
            }

            var boundary = inLeft ? piece : piece.Reversed();
            if (!kept.Any(k => k.IsSameAs(boundary, tolerance)))
            {
                kept.Add(boundary);
            }
        }

        result.Pieces = kept;
        result.Loops = Chain(kept, tolerance);
        result.Measure();
        return result;
    }

    /// <summary>The tolerance for two places to be one, at the scale of the coordinates: a touch by construction is a hair apart</summary>
    static double Tolerance(double scale)
    {
        return 1e-9 * scale;
    }

    static double ScaleOf(IEnumerable<BoundaryPiece> pieces)
    {
        double result = 1;
        foreach (var piece in pieces)
        {
            var bounds = piece.Bounds;
            result = System.Math.Max(result, System.Math.Max(
                System.Math.Max(System.Math.Abs(bounds.Left), System.Math.Abs(bounds.Right)),
                System.Math.Max(System.Math.Abs(bounds.Top), System.Math.Abs(bounds.Bottom))));
        }

        return result;
    }

    static bool IsInside(IRegion operand, Point point)
    {
        return operand.HasInside && operand.ContainsPoint(point);
    }

    /// <summary>
    /// The piece with its operand on the left, or null when the operand is on neither side
    /// or on both: judged a hair off the middle of the piece, inside and out - the operand's
    /// own boundary is the piece there, and its other pieces are at least the piece's length
    /// away
    /// </summary>
    static BoundaryPiece OrientToOperand(BoundaryPiece piece, IRegion operand, double scale)
    {
        double length = piece.Length;
        double step = System.Math.Max(1e-13 * scale, System.Math.Min(1e-7 * scale, length / 100));
        var middle = piece.Middle;
        var normal = piece.LeftNormalAt(0.5);
        bool left = IsInside(operand, new Point(middle.X + normal.X * step, middle.Y + normal.Y * step));
        bool right = IsInside(operand, new Point(middle.X - normal.X * step, middle.Y - normal.Y * step));
        if (left == right)
        {
            return null;
        }

        return left ? piece : piece.Reversed();
    }

    #endregion

    #region The operands' boundaries

    /// <summary>At most this much of a turn in one arc piece: a drawn arc can't be a whole turn, and the crossings are found along shorter arcs</summary>
    const double MaxArcSweep = System.Math.PI / 2;

    /// <summary>The boundary of a figure with an inside, in pieces, in any order and either way round</summary>
    public static List<BoundaryPiece> PiecesOf(IRegion figure)
    {
        var result = new List<BoundaryPiece>();
        switch (figure)
        {
            case Region region:
                result.AddRange(region.Boundary.Pieces);
                break;
            case PolygonBase polygon:
                AddPolygon(result, polygon.CurrentVertexCoordinates());
                break;
            case EllipseArcBase arc when figure is IShapeWithInterior:
                {
                    var (begin, sweep) = arc.ParametricArc;
                    AddArc(result, arc.Center, arc.SemiMajor, arc.SemiMinor, arc.Inclination, begin, begin + sweep);
                    var start = arc.BeginLocation;
                    var end = arc.EndLocation;
                    if (figure is CircleSector || figure is EllipseSector)
                    {
                        var center = arc.Center;
                        result.Add(new SegmentPiece(end, center));
                        result.Add(new SegmentPiece(center, start));
                    }
                    else
                    {
                        result.Add(new SegmentPiece(end, start));
                    }

                    break;
                }

            case EllipseBase ellipse:
                AddArc(result, ellipse.Center, ellipse.SemiMajor, ellipse.SemiMinor, ellipse.Inclination, 0, Math.DOUBLEPI);
                break;
            case BezierPath path:
                AddPolygon(result, path.OutlinePolygon());
                foreach (var hole in path.Holes)
                {
                    if (hole.Exists && hole.PieceCount > 0)
                    {
                        AddPolygon(result, hole.OutlinePolygon());
                    }
                }

                break;
        }

        return result.Where(piece => piece.Length > 0 && piece.Start.Exists() && piece.End.Exists()).ToList();
    }

    static void AddPolygon(List<BoundaryPiece> result, IList<Point> vertices)
    {
        int count = vertices.Count;
        if (count < 3)
        {
            return;
        }

        for (int i = 0; i < count; i++)
        {
            var from = vertices[i];
            var to = vertices[(i + 1) % count];
            if (from != to)
            {
                result.Add(new SegmentPiece(from, to));
            }
        }
    }

    /// <summary>An arc in pieces of at most a quarter turn each</summary>
    static void AddArc(
        List<BoundaryPiece> result,
        Point center,
        double semiMajor,
        double semiMinor,
        double inclination,
        double from,
        double to)
    {
        if (!(semiMajor > 0 && semiMinor > 0) || !(System.Math.Abs(to - from) > 0))
        {
            return;
        }

        int count = (int)System.Math.Ceiling(System.Math.Abs(to - from) / MaxArcSweep - 1e-9);
        count = System.Math.Max(1, count);
        for (int i = 0; i < count; i++)
        {
            double a = from + (to - from) * i / count;
            double b = from + (to - from) * (i + 1) / count;
            result.Add(new ArcPiece(center, semiMajor, semiMinor, inclination, a, b));
        }
    }

    #endregion

    #region Cutting at the crossings

    static List<(BoundaryPiece Piece, int Operand)> Split(List<(BoundaryPiece Piece, int Operand)> pieces, double tolerance)
    {
        var cuts = new List<double>[pieces.Count];
        for (int i = 0; i < pieces.Count; i++)
        {
            cuts[i] = new List<double>();
        }

        var bounds = pieces.Select(p => p.Piece.Bounds.Inflate(tolerance)).ToArray();
        for (int i = 0; i < pieces.Count; i++)
        {
            for (int j = i + 1; j < pieces.Count; j++)
            {
                if (bounds[i].Intersects(bounds[j]))
                {
                    FindCrossings(pieces[i].Piece, pieces[j].Piece, cuts[i], cuts[j], tolerance);
                }
            }
        }

        var result = new List<(BoundaryPiece, int)>();
        for (int i = 0; i < pieces.Count; i++)
        {
            var piece = pieces[i].Piece;
            double parameterTolerance = tolerance / piece.Length;
            var parameters = cuts[i]
                .Where(t => t > parameterTolerance && t < 1 - parameterTolerance)
                .OrderBy(t => t)
                .ToList();
            double last = 0;
            foreach (var t in parameters)
            {
                if (t - last > parameterTolerance)
                {
                    result.Add((piece.Part(last, t), pieces[i].Operand));
                    last = t;
                }
            }

            result.Add((last == 0 ? piece : piece.Part(last, 1), pieces[i].Operand));
        }

        return result;
    }

    /// <summary>The parameters where the two pieces cross, each added to its own piece's list (where a crossing lies past an end, nothing)</summary>
    static void FindCrossings(
        BoundaryPiece p,
        BoundaryPiece q,
        List<double> cutsP,
        List<double> cutsQ,
        double tolerance)
    {
        switch (p, q)
        {
            case (SegmentPiece sp, SegmentPiece sq):
                CrossSegments(sp, sq, cutsP, cutsQ, tolerance);
                break;
            case (SegmentPiece segment, ArcPiece arc):
                CrossSegmentAndArc(segment, arc, cutsP, cutsQ, tolerance);
                break;
            case (ArcPiece arc, SegmentPiece segment):
                CrossSegmentAndArc(segment, arc, cutsQ, cutsP, tolerance);
                break;
            case (ArcPiece ap, ArcPiece aq):
                CrossArcs(ap, aq, cutsP, cutsQ, tolerance);
                break;
        }
    }

    static double Cross(Point a, Point b)
    {
        return a.X * b.Y - a.Y * b.X;
    }

    static void CrossSegments(
        SegmentPiece p,
        SegmentPiece q,
        List<double> cutsP,
        List<double> cutsQ,
        double tolerance)
    {
        var r = p.End.Minus(p.Start);
        var s = q.End.Minus(q.Start);
        var offset = q.Start.Minus(p.Start);
        double lengthP = p.Length;
        double lengthQ = q.Length;
        double denominator = Cross(r, s);
        if (System.Math.Abs(denominator) <= 1e-12 * lengthP * lengthQ)
        {
            // parallel: on one line, each is cut where the other's ends lie on it
            if (System.Math.Abs(Cross(offset, r)) <= tolerance * lengthP)
            {
                AddWithin(cutsP, p.ParameterOf(q.Start), tolerance / lengthP);
                AddWithin(cutsP, p.ParameterOf(q.End), tolerance / lengthP);
                AddWithin(cutsQ, q.ParameterOf(p.Start), tolerance / lengthQ);
                AddWithin(cutsQ, q.ParameterOf(p.End), tolerance / lengthQ);
            }

            return;
        }

        double t = Cross(offset, s) / denominator;
        double u = Cross(offset, r) / denominator;
        if (IsWithin(t, tolerance / lengthP) && IsWithin(u, tolerance / lengthQ))
        {
            cutsP.Add(t);
            cutsQ.Add(u);
        }
    }

    static bool IsWithin(double t, double parameterTolerance)
    {
        return t >= -parameterTolerance && t <= 1 + parameterTolerance;
    }

    static void AddWithin(List<double> cuts, double t, double parameterTolerance)
    {
        if (IsWithin(t, parameterTolerance))
        {
            cuts.Add(t);
        }
    }

    static void CrossSegmentAndArc(
        SegmentPiece segment,
        ArcPiece arc,
        List<double> cutsSegment,
        List<double> cutsArc,
        double tolerance)
    {
        var crossings = Math.GetIntersectionOfEllipseAndLine(
            arc.Center,
            arc.SemiMajor,
            arc.SemiMinor,
            arc.Inclination,
            new PointPair(segment.Start, segment.End));
        AddCrossing(segment, arc, crossings.P1, cutsSegment, cutsArc, tolerance);
        if (!Near(crossings.P1, crossings.P2, tolerance))
        {
            AddCrossing(segment, arc, crossings.P2, cutsSegment, cutsArc, tolerance);
        }
    }

    static void AddCrossing(
        SegmentPiece segment,
        ArcPiece arc,
        Point point,
        List<double> cutsSegment,
        List<double> cutsArc,
        double tolerance)
    {
        if (!point.Exists())
        {
            return;
        }

        double t = segment.ParameterOf(point);
        double u = arc.ParameterOf(arc.AngleOf(point), AngularTolerance(arc, tolerance));
        if (IsWithin(t, tolerance / segment.Length) && !double.IsNaN(u))
        {
            cutsSegment.Add(t);
            cutsArc.Add(u);
        }
    }

    static double AngularTolerance(ArcPiece arc, double tolerance)
    {
        return tolerance / System.Math.Max(arc.SemiMajor, arc.SemiMinor);
    }

    static void CrossArcs(
        ArcPiece p,
        ArcPiece q,
        List<double> cutsP,
        List<double> cutsQ,
        double tolerance)
    {
        bool sameEllipse = Near(p.Center, q.Center, tolerance)
            && System.Math.Abs(p.SemiMajor - q.SemiMajor) <= tolerance
            && System.Math.Abs(p.SemiMinor - q.SemiMinor) <= tolerance
            && (p.IsCircle && q.IsCircle || System.Math.Abs(System.Math.IEEERemainder(p.Inclination - q.Inclination, Math.PI)) <= 1e-9);
        if (sameEllipse)
        {
            // on one curve: each is cut where the other's ends lie on it
            AddArcCut(cutsP, p, q.Start, tolerance);
            AddArcCut(cutsP, p, q.End, tolerance);
            AddArcCut(cutsQ, q, p.Start, tolerance);
            AddArcCut(cutsQ, q, p.End, tolerance);
            return;
        }

        if (p.IsCircle && q.IsCircle)
        {
            var crossings = Math.GetIntersectionOfCircles(p.Center, p.SemiMajor, q.Center, q.SemiMajor);
            AddArcCrossing(p, q, crossings.P1, cutsP, cutsQ, tolerance);
            if (!Near(crossings.P1, crossings.P2, tolerance))
            {
                AddArcCrossing(p, q, crossings.P2, cutsP, cutsQ, tolerance);
            }

            return;
        }

        // an ellipse among them: the roots of its equation along the other arc, by sampling
        // and bisection (a touch, where the equation only grazes 0, is no crossing)
        const int steps = 48;
        double previous = p.Implicit(q.PointAt(0));
        for (int i = 1; i <= steps; i++)
        {
            double t = (double)i / steps;
            double value = p.Implicit(q.PointAt(t));
            if (previous == 0 || previous * value < 0)
            {
                double a = (double)(i - 1) / steps;
                double b = t;
                double valueA = previous;
                if (previous != 0)
                {
                    for (int k = 0; k < 60; k++)
                    {
                        double m = (a + b) / 2;
                        double valueM = p.Implicit(q.PointAt(m));
                        if (valueM == 0)
                        {
                            a = b = m;
                            break;
                        }

                        if (valueA * valueM < 0)
                        {
                            b = m;
                        }
                        else
                        {
                            a = m;
                            valueA = valueM;
                        }
                    }
                }
                else
                {
                    b = a;
                }

                double root = (a + b) / 2;
                AddArcCrossing(p, q, q.PointAt(root), cutsP, cutsQ, tolerance);
            }

            previous = value;
        }
    }

    static void AddArcCut(List<double> cuts, ArcPiece arc, Point point, double tolerance)
    {
        double t = arc.ParameterOf(arc.AngleOf(point), AngularTolerance(arc, tolerance));
        if (!double.IsNaN(t))
        {
            cuts.Add(t);
        }
    }

    static void AddArcCrossing(
        ArcPiece p,
        ArcPiece q,
        Point point,
        List<double> cutsP,
        List<double> cutsQ,
        double tolerance)
    {
        if (!point.Exists())
        {
            return;
        }

        double t = p.ParameterOf(p.AngleOf(point), AngularTolerance(p, tolerance));
        double u = q.ParameterOf(q.AngleOf(point), AngularTolerance(q, tolerance));
        if (!double.IsNaN(t) && !double.IsNaN(u))
        {
            cutsP.Add(t);
            cutsQ.Add(u);
        }
    }

    #endregion

    #region Loops and measures

    /// <summary>
    /// The pieces chained end to start into closed loops. Where several meet at one point
    /// (a touch) any pairing will do: what is filled goes by the pieces, not by how they
    /// are chained. A piece whose end nothing continues closes its loop where it is.
    /// </summary>
    static IReadOnlyList<IReadOnlyList<BoundaryPiece>> Chain(List<BoundaryPiece> pieces, double tolerance)
    {
        var result = new List<IReadOnlyList<BoundaryPiece>>();
        var used = new bool[pieces.Count];
        for (int start = 0; start < pieces.Count; start++)
        {
            if (used[start])
            {
                continue;
            }

            var loop = new List<BoundaryPiece>() { pieces[start] };
            used[start] = true;
            var origin = pieces[start].Start;
            var end = pieces[start].End;
            while (!Near(end, origin, tolerance))
            {
                int next = -1;
                double best = double.PositiveInfinity;
                for (int i = 0; i < pieces.Count; i++)
                {
                    if (used[i])
                    {
                        continue;
                    }

                    double distance = pieces[i].Start.Distance(end);
                    if (distance <= tolerance && distance < best)
                    {
                        best = distance;
                        next = i;
                    }
                }

                if (next < 0)
                {
                    break;
                }

                used[next] = true;
                loop.Add(pieces[next]);
                end = pieces[next].End;
            }

            result.Add(loop);
        }

        return result;
    }

    void Measure()
    {
        if (Pieces.Count == 0)
        {
            Area = 0;
            Perimeter = 0;
            Centroid = Math.InfinitePoint;
            return;
        }

        var origin = Pieces[0].Start;
        double area = 0;
        double perimeter = 0;
        foreach (var piece in Pieces)
        {
            area += piece.AreaTerm(origin);
            perimeter += piece.Length;
        }

        Area = area;
        Perimeter = perimeter;
        Centroid = FindCentroid(origin);
    }

    /// <summary>The centroid of the loops as polygons through their samples (holes count against), or the middle of the pieces' ends where the area is nothing</summary>
    Point FindCentroid(Point origin)
    {
        double area = 0;
        double x = 0;
        double y = 0;
        foreach (var loop in Loops)
        {
            var points = new List<Point>() { loop[0].Start.Minus(origin) };
            foreach (var piece in loop)
            {
                var samples = new List<Point>();
                piece.AddSamples(samples);
                points.AddRange(samples.Select(sample => sample.Minus(origin)));
            }

            for (int i = 0; i < points.Count; i++)
            {
                var a = points[i];
                var b = points[(i + 1) % points.Count];
                double cross = a.X * b.Y - b.X * a.Y;
                area += cross;
                x += (a.X + b.X) * cross;
                y += (a.Y + b.Y) * cross;
            }
        }

        if (System.Math.Abs(area) > 1e-12 * scale * scale)
        {
            return new Point(origin.X + x / (3 * area), origin.Y + y / (3 * area));
        }

        var ends = Pieces.Select(piece => piece.Start).ToList();
        return new Point(ends.Average(p => p.X), ends.Average(p => p.Y));
    }

    #endregion

    #region Drawing

    /// <summary>The loops as a geometry in pixels, filled even-odd; one empty figure where there is nothing, so that the path repaints</summary>
    public PathGeometry CreateGeometry(CoordinateSystem coordinateSystem)
    {
        var geometry = new PathGeometry() { FillRule = FillRule.EvenOdd };
        foreach (var loop in Loops)
        {
            var figure = new PathFigure()
            {
                StartPoint = coordinateSystem.ToPhysical(loop[0].Start),
                IsClosed = true,
                IsFilled = true,
                Segments = new PathSegmentCollection()
            };
            foreach (var piece in loop)
            {
                figure.Segments.Add(ToSegment(piece, coordinateSystem));
            }

            geometry.Figures.Add(figure);
        }

        if (geometry.Figures.Count == 0)
        {
            geometry.Figures.Add(new PathFigure()
            {
                IsClosed = false,
                IsFilled = false,
                Segments = new PathSegmentCollection()
            });
        }

        return geometry;
    }

    static PathSegment ToSegment(BoundaryPiece piece, CoordinateSystem coordinateSystem)
    {
        if (piece is ArcPiece arc)
        {
            return new ArcSegment()
            {
                Point = coordinateSystem.ToPhysical(arc.End),
                Size = new Size(coordinateSystem.ToPhysical(arc.SemiMajor), coordinateSystem.ToPhysical(arc.SemiMinor)),
                RotationAngle = -arc.Inclination.ToDegrees(),
                IsLargeArc = System.Math.Abs(arc.Sweep) > Math.PI,
                SweepDirection = arc.IsCounterclockwise ? SweepDirection.CounterClockwise : SweepDirection.Clockwise
            };
        }

        return new LineSegment() { Point = coordinateSystem.ToPhysical(piece.End) };
    }

    #endregion
}
