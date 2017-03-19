using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using M = System.Math;

namespace DynamicGeometry
{
    public struct PointPair
    {
        public PointPair(Point p1, Point p2)
        {
            P1 = p1;
            P2 = p2;
        }

        public Point P1;
        public Point P2;

        public bool Contains(Point point)
        {
            return point.X >= P1.X
                && point.X <= P2.X
                && point.Y >= P1.Y
                && point.Y <= P2.Y;
        }

        public bool ContainsInner(Point point)
        {
            return point.X > P1.X
                && point.X < P2.X
                && point.Y > P1.Y
                && point.Y < P2.Y;
        }

        public PointPair Reverse
        {
            get
            {
                return new PointPair(P2, P1);
            }
        }

        public PointPair GetBoundingRect()
        {
            PointPair result = this;
            if (result.P1.X > result.P2.X)
            {
                var t = result.P1.X;
                result.P1.X = result.P2.X;
                result.P2.X = t;
            }
            if (result.P1.Y > result.P2.Y)
            {
                var t = result.P1.Y;
                result.P1.Y = result.P2.Y;
                result.P2.Y = t;
            }
            return result;
        }

        public double Length
        {
            get
            {
                return P1.Distance(P2);
            }
        }

        public Point Midpoint
        {
            get
            {
                return new Point((P1.X + P2.X) / 2, (P1.Y + P2.Y) / 2);
            }
        }

        public override string ToString()
        {
            return string.Format("({0};{1})-({2};{3})", P1.X, P1.Y, P2.X, P2.Y);
        }
    }

    public static class PointExtensions
    {
        public static double Distance(this Point p1, Point p2)
        {
            return (
                  (p1.X - p2.X).Sqr()
                + (p1.Y - p2.Y).Sqr()).SquareRoot();
        }

        public static double Length(this Point p)
        {
            return p.SumOfSquares().SquareRoot();
        }

        public static Point Scale(this Point p, double scaleFactor)
        {
            return new Point(p.X * scaleFactor, p.Y * scaleFactor);
        }

        public static Point TrimToMaxLength(this Point p, double maxLength)
        {
            var length = p.Length();
            if (maxLength > 0 && length > maxLength)
            {
                var ratio = maxLength / length;
                return new Point(p.X * ratio, p.Y * ratio);
            }
            return p;
        }

        public static Point Minus(this Point point)
        {
            return new Point(-point.X, -point.Y);
        }

        public static Point Minus(this Point point, Point other)
        {
            return new Point(point.X - other.X, point.Y - other.Y);
        }

        public static Point OffsetX(this Point point, double xOffset)
        {
            return new Point(point.X + xOffset, point.Y);
        }

        public static Point OffsetY(this Point point, double yOffset)
        {
            return new Point(point.X, point.Y + yOffset);
        }

        public static Point Offset(this Point point, double xOffset, double yOffset)
        {
            return new Point(point.X + xOffset, point.Y + yOffset);
        }

        public static Point Plus(this Point point, Point other)
        {
            return new Point(point.X + other.X, point.Y + other.Y);
        }

        public static double SumOfSquares(this Point point)
        {
            return point.X.Sqr() + point.Y.Sqr();
        }

        public static bool Exists(this Point p)
        {
            return !double.IsNaN(p.X) && !double.IsNaN(p.Y);
        }

        public static bool IsValidPositiveValue(this double value)
        {
            return value.IsValidValue()
                && value > 0;
        }

        public static bool IsValidNonNegativeValue(this double value)
        {
            return value.IsValidValue()
                && value >= 0;
        }

        public static bool IsValidValue(this double value)
        {
            return !double.IsNaN(value)
                && !double.IsInfinity(value);
        }

        public static bool IsWithinEpsilonTo(this double value, double center)
        {
            return Math.Abs(value - center) < Math.Epsilon;
        }

        public static bool IsWithinEpsilon(this double value)
        {
            return Math.Abs(value) < Math.Epsilon;
        }

        public static bool IsWithinTolerance(this double value)
        {
            return Math.Abs(value) < Math.CursorTolerance;
        }
    }

    public static class Math
    {
        public static Point InfinitePoint
        {
            get
            {
                return new Point(double.NaN, double.NaN);
            }
        }

        public static PointPair InfinitePointPair
        {
            get
            {
                return new PointPair() { P1 = Math.InfinitePoint, P2 = Math.InfinitePoint };
            }
        }

        public static bool IsPointInPolygon(this IList<Point> polygon, Point start)
        {
            var outsidePoint = new Point(polygon.Max(p => p.X) + 100, 0);
            var segment = new PointPair(start, outsidePoint);

            int n = polygon.Count;
            int intersectionsCount = 0;
            for (int i = 0; i < n; i++)
            {
                PointPair side = new PointPair(polygon[i], polygon[i.RotateNext(n)]);

                Point intersectionPoint = GetIntersectionOfLines(segment, side);
                if (!intersectionPoint.Exists())
                {
                    continue;
                }
                if (IsPointInSegmentBoundingRect(side, intersectionPoint)
                    && IsPointInSegmentBoundingRect(segment, intersectionPoint)
                    && intersectionPoint != side.P2)
                {
                    if (intersectionPoint == side.P1 
                        && Math.VectorProduct(start, side.P1, polygon[i.RotatePrevious(n)]).Sign()
                        == Math.VectorProduct(start, side.P1, side.P2).Sign())
                    {
                        continue;
                    }
                    intersectionsCount++;
                }
            }

            return intersectionsCount % 2 == 1;
        }

        public static List<Point> GetIntersections(IList<Point> polygon, PointPair segment)
        {
            List<Point> result = new List<Point>();
            int n = polygon.Count;
            for (int i = 0; i < n; i++)
            {
                PointPair side = new PointPair(polygon[i], polygon[i.RotateNext(n)]);
                Point intersection = GetIntersectionOfSegments(side, segment);
                if (intersection.Exists() && intersection != polygon[i.RotateNext(n)])
                {
                    result.Add(intersection);
                }
            }
            return result;
        }

        public static double PolylineLength(this IList<Point> polyline)
        {
            double sum = 0;
            for (int i = 0; i < polyline.Count - 1; i++)
            {
                sum += polyline[i].Distance(polyline[i + 1]);
            }
            return sum;
        }

        public static int Sign(this double num)
        {
            return M.Sign(num);
        }

        public static double Abs(this double num)
        {
            return M.Abs(num);
        }

        public static double Round(this double num, int fractionalDigits)
        {
            return M.Round(num, fractionalDigits);
        }

        private static double mCursorTolerance = 5;
        public static double CursorTolerance
        {
            get
            {
                return mCursorTolerance;
            }
            set
            {
                mCursorTolerance = value;
            }
        }

        public static Point ScalePointBetweenTwo(Point p1, Point p2, double ratio)
        {
            return new Point(
                p1.X + (p2.X - p1.X) * ratio,
                p1.Y + (p2.Y - p1.Y) * ratio);
        }

        public static double OAngle(Point firstPoint, Point vertex, Point secondPoint)
        {
            var a1 = GetAngle(vertex, firstPoint);
            var a2 = GetAngle(vertex, secondPoint);
            if (a2 < a1)
            {
                a2 = a2 + 2 * PI;
            }
            var result = 2 * PI + a1 - a2;
            if (a1 == a2)
            {
                result = 0;
            }
            return result;
        }

        public static double GetAngle(Point center, Point endPoint)
        {
            return M.Atan2(endPoint.Y - center.Y, endPoint.X - center.X) + PI;
        }

        public static double VectorProduct(Point p1, Point p2, Point p3)
        {
            return (p2.X - p1.X) * (p3.Y - p1.Y) - (p3.X - p1.X) * (p2.Y - p1.Y);
        }

        public static double PI
        {
            get
            {
                return M.PI;
            }
        }

        /// <summary>
        /// Square of a number (multiplied by itself)
        /// </summary>
        public static double Sqr(this double num)
        {
            return num * num;
        }

        /// <summary>
        /// Square root
        /// </summary>
        public static double SquareRoot(this double num)
        {
            return System.Math.Sqrt(System.Math.Abs(num));
        }

        public static void Swap<T>(ref T p1, ref T p2)
        {
            T temp = p1;
            p1 = p2;
            p2 = temp;
        }

        public static PointPair GetLineFromSegment(PointPair segment, PointPair borders)
        {
            PointPair result = new PointPair();

            if (segment.P1 == segment.P2)
            {
                result = segment;
            }
            else if (segment.P1.X == segment.P2.X)
            {
                result.P1.X = segment.P1.X;
                result.P2.X = segment.P1.X;
                result.P1.Y = borders.P1.Y;
                result.P2.Y = borders.P2.Y;
                if (segment.P1.Y > segment.P2.Y)
                {
                    result.P1.Y = borders.P2.Y;
                    result.P2.Y = borders.P1.Y;
                }
            }
            else if (segment.P1.Y == segment.P2.Y)
            {
                result.P1.X = borders.P1.X;
                result.P1.Y = segment.P1.Y;
                result.P2.X = borders.P2.X;
                result.P2.Y = segment.P1.Y;
                if (segment.P1.X > segment.P2.X)
                {
                    result.P1.X = borders.P2.X;
                    result.P2.X = borders.P1.X;
                }
            }
            else
            {
                var deltaX = segment.P2.X - segment.P1.X;
                var deltaY = segment.P2.Y - segment.P1.Y;
                var deltaXYRatio = deltaX / deltaY;
                var deltaYXRatio = deltaY / deltaX;

                result.P1.Y = deltaY > 0 ? borders.P1.Y : borders.P2.Y;
                result.P1.X = segment.P1.X + (result.P1.Y - segment.P1.Y) * deltaXYRatio;
                if (result.P1.X < borders.P1.X)
                {
                    result.P1.X = borders.P1.X;
                    result.P1.Y = segment.P1.Y + (result.P1.X - segment.P1.X) * deltaYXRatio;
                }
                else if (result.P1.X > borders.P2.X)
                {
                    result.P1.X = borders.P2.X;
                    result.P1.Y = segment.P1.Y + (result.P1.X - segment.P1.X) * deltaYXRatio;
                }

                result.P2.X = deltaX > 0 ? borders.P2.X : borders.P1.X;
                result.P2.Y = segment.P2.Y + (result.P2.X - segment.P2.X) * deltaYXRatio;
                if (result.P2.Y < borders.P1.Y)
                {
                    result.P2.Y = borders.P1.Y;
                    result.P2.X = segment.P2.X + (result.P2.Y - segment.P2.Y) * deltaXYRatio;
                }
                else if (result.P2.Y > borders.P2.Y)
                {
                    result.P2.Y = borders.P2.Y;
                    result.P2.X = segment.P2.X + (result.P2.Y - segment.P2.Y) * deltaXYRatio;
                }
            }

            return result;
        }

        public class ProjectionInfo
        {
            public Point Point { get; set; }
            public double Ratio { get; set; }
            public double DistanceToLine { get; set; }
        }

        public static ProjectionInfo GetProjection(Point point, PointPair line)
        {
            Point projectionPoint = GetProjectionPoint(point, line);
            ProjectionInfo result = new ProjectionInfo()
            {
                Point = projectionPoint,
                Ratio = GetProjectionRatio(line, projectionPoint),
                DistanceToLine = projectionPoint.Distance(point)
            };
            return result;
        }

        public static double GetProjectionRatio(PointPair line, Point projection)
        {
            var result = 0.0;
            if (line.P1.X != line.P2.X)
            {
                result = (projection.X - line.P1.X) / (line.P2.X - line.P1.X);
            }
            else if (line.P1.Y != line.P2.Y)
            {
                result = (projection.Y - line.P1.Y) / (line.P2.Y - line.P1.Y);
            }
            return result;
        }

        public static Point GetProjectionPoint(Point p, PointPair line)
        {
            Point result = new Point();

            if (line.P1.Y == line.P2.Y)
            {
                result.X = p.X;
                result.Y = line.P1.Y;
            }
            else if (line.P1.X == line.P2.X)
            {
                result.X = line.P1.X;
                result.Y = p.Y;
            }
            else
            {
                var a = p.Minus(line.P1).SumOfSquares();
                var b = p.Minus(line.P2).SumOfSquares();
                var c = line.P1.Minus(line.P2).SumOfSquares();

                if (c != 0)
                {
                    var m = (a + c - b) / (2 * c);
                    result = ScalePointBetweenTwo(line.P1, line.P2, m);
                }
                else
                {
                    result = line.P1;
                }
            }

            return result;
        }

        public static Point GetIntersectionOfSegments(PointPair segment1, PointPair segment2)
        {
            Point result = GetIntersectionOfLines(segment1, segment2);
            if (!result.Exists())
            {
                return result;
            }
            if (IsPointInSegmentInnerBoundingRect(segment1, result)
                && IsPointInSegmentInnerBoundingRect(segment2, result))
            {
                return result;
            }
            return Math.InfinitePoint;
        }

        public static bool IsPointInSegmentBoundingRect(PointPair segment, Point point)
        {
            var boundingRect = segment.GetBoundingRect();
            return boundingRect.Contains(point);
        }

        public static bool IsPointInSegmentInnerBoundingRect(PointPair segment, Point point)
        {
            segment = segment.GetBoundingRect();
            if (segment.P1.X == segment.P2.X)
            {
                return point.X == segment.P1.X
                    && point.Y > segment.P1.Y
                    && point.Y < segment.P2.Y;
            }
            else if (segment.P1.Y == segment.P2.Y)
            {
                return point.Y == segment.P1.Y
                    && point.X > segment.P1.X
                    && point.X < segment.P2.X;
            }
            return segment.ContainsInner(point);
        }

        public static Point GetIntersectionOfLines(PointPair line1, PointPair line2)
        {
            var a1 = line1.P2.Y - line1.P1.Y;
            var b1 = line1.P1.X - line1.P2.X;
            var c1 = line1.P2.X * line1.P1.Y - line1.P1.X * line1.P2.Y;

            var a2 = line2.P2.Y - line2.P1.Y;
            var b2 = line2.P1.X - line2.P2.X;
            var c2 = line2.P2.X * line2.P1.Y - line2.P1.X * line2.P2.Y;

            return SolveLinearSystem(a1, b1, c1, a2, b2, c2);
        }

        public static Point SolveLinearSystem(
            double a1, double b1, double c1,
            double a2, double b2, double c2)
        {
            var d = a1 * b2 - a2 * b1;
            if (d == 0)
            {
                return Math.InfinitePoint;
            }

            var dx = b1 * c2 - b2 * c1;
            var dy = a2 * c1 - a1 * c2;
            return new Point(dx / d, dy / d);
        }

        public static ReadOnlyCollection<double> SolveSquareEquation(
            double a,
            double b,
            double c)
        {
            var result = new List<double>(2);
            var d = b * b - 4 * a * c;
            if (a == 0)
            {
                return new ReadOnlyCollection<double>(result);
            }

            if (d > 0)
            {
                d = d.SquareRoot();
                a *= 2;
                result.Add((-b - d) / a);
                result.Add((d - b) / a);
            }
            else if (d == 0)
            {
                result.Add(-b / (2 * a));
            }
            return new ReadOnlyCollection<double>(result);
        }

        public static Point Midpoint(this IEnumerable<Point> points)
        {
            if (!points.Any())
            {
                return new Point();
            }
            Point sum = new Point();
            foreach (var point in points)
            {
                sum = sum.Plus(point);
            }
            return sum.Scale(1.0 / points.Count());
        }

        public static Point Midpoint(params Point[] points)
        {
            return points.Midpoint();
        }

        public static double Area(this IEnumerable<Point> vertices)
        {
            var points = vertices.ToArray();
            if (points.Length < 2)
            {
                return 0;
            }
            else if (points.Length == 2)
            {
                // if only two points are given, assume area of a circle
                // with center in the first point and passing through the second one
                return points[0].Distance(points[1]).Sqr() * PI;
            }
            else
            {
                // general polygon area
                double sum = 0;
                for (int i = 0; i < points.Length - 1; i++)
                {
                    sum += (points[i + 1].X - points[i].X) * (points[i + 1].Y + points[i].Y) / 2;
                }
                var lastIndex = points.Length - 1;
                sum += (points[0].X - points[lastIndex].X) * (points[0].Y + points[lastIndex].Y) / 2;
                return sum.Abs();
            }
        }

        public static double Epsilon
        {
            get
            {
                return 0.1;
            }
        }

        public static PointPair GetIntersectionOfCircleAndLine(
            Point center,
            double radius,
            PointPair line)
        {
            var result = Math.InfinitePointPair;
            var p = GetProjectionPoint(center, line);
            var h = center.Distance(p).Round(4);
            radius = radius.Round(4);

            if ((h - radius).Abs() < Math.Epsilon)
            {
                result.P1 = p;
                result.P2 = p;
            }
            else if (h > 0 && h < radius)
            {
                var s = ((radius - h) * (radius + h)).SquareRoot();
                s = s / line.Length;
                result.P1.X = p.X - (line.P2.X - line.P1.X) * s;
                result.P1.Y = p.Y - (line.P2.Y - line.P1.Y) * s;
                result.P2.X = 2 * p.X - result.P1.X;
                result.P2.Y = 2 * p.Y - result.P1.Y;
            }
            else if (h == 0)
            {
                var a = line.P1.Distance(line.P2);
                if (a != 0)
                {
                    var s = radius / a;
                    result.P1.X = center.X + (line.P2.X - line.P1.X) * s;
                    result.P1.Y = center.Y + (line.P2.Y - line.P1.Y) * s;
                    result.P2.X = 2 * center.X - result.P1.X;
                    result.P2.Y = 2 * center.Y - result.P1.Y;
                }
            }

            return result;
        }

        public static PointPair GetIntersectionOfCircles(
            Point center1,
            double radius1,
            Point center2,
            double radius2)
        {
            var result = Math.InfinitePointPair;
            var x1 = center1.X;
            var y1 = center1.Y;
            var x2 = center2.X;
            var y2 = center2.Y;

            var r3 = center1.Distance(center2);
            if (r3 == 0)
            {
                return result;
            }

            if (y1 == y2)
            {
                if (radius1 + radius1 > r3 + Epsilon
                 && radius1 + r3 > radius2 + Epsilon
                 && radius2 + r3 > radius1 + Epsilon
                 && x1 != x2)
                {
                    var x3 = (radius1.Sqr() + r3.Sqr() - radius2.Sqr()) / (2 * r3);
                    var tsqr = (radius1.Sqr() - x3.Sqr()).SquareRoot();
                    result.P1.X = x1 + (x2 - x1) * x3 / r3;
                    result.P1.Y = y1 - tsqr;
                    result.P2.X = result.P1.X;
                    result.P2.Y = y1 + tsqr;
                    if (x2 < x1)
                    {
                        var t = result.P1.Y;
                        result.P1.Y = result.P2.Y;
                        result.P2.Y = t;
                    }
                }
                else if ((radius1 + radius2 - r3).Abs() <= Epsilon)
                {
                    result.P1.X = x1 + M.Sign(x2 - x1) * radius1;
                    result.P1.Y = y1;
                    result.P2 = result.P1;
                }
                else if ((radius1 + r3 - radius2).Abs() <= Epsilon
                    || (r3 + radius2 - radius1).Abs() <= Epsilon)
                {
                    result.P1.X = x1 + M.Sign(x2 - x1) * M.Sign(radius1 - radius2) * radius1;
                    result.P1.Y = y1;
                    result.P2 = result.P1;
                }
                return result;
            }

            if ((radius1 + radius2 - r3).Abs() <= Epsilon)
            {
                r3 = radius1 / r3;
                result.P1.X = x1 + (x2 - x1) * r3;
                result.P1.Y = y1 + (y2 - y1) * r3;
                result.P2 = result.P1;
                return result;
            }

            if ((radius1 + r3 - radius2).Abs() <= Epsilon
                || (radius2 + r3 - radius1).Abs() <= Epsilon)
            {
                r3 = radius1 / r3 * M.Sign(radius1 - radius2);
                result.P1.X = x1 + (x2 - x1) * r3;
                result.P1.Y = y1 + (y2 - y1) * r3;
                result.P2 = result.P1;
                return result;
            }

            var k = -(x2 - x1) / (y2 - y1);
            var b = ((radius1 - radius2) * (radius1 + radius2)
                    + (x2 - x1) * (x2 + x1)
                    + (y2 - y1) * (y2 + y1))
                / (2 * (y2 - y1));
            var ea = k * k + 1;
            var eb = 2 * (k * b - x1 - k * y1);
            var ec = x1.Sqr() + b.Sqr() - 2 * b * y1 + y1.Sqr() - radius1.Sqr();
            var roots = SolveSquareEquation(ea, eb, ec);
            if (roots == null || roots.Count != 2)
            {
                return result;
            }
            result.P1.X = roots[0];
            result.P1.Y = roots[0] * k + b;
            result.P2.X = roots[1];
            result.P2.Y = roots[1] * k + b;

            if (y2 > y1)
            {
                var t = result.P1;
                result.P1 = result.P2;
                result.P2 = t;
            }

            return result;
        }
    }
}