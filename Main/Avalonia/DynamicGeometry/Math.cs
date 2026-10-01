using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Avalonia;
using M = System.Math;

namespace DynamicGeometry
{

    public partial struct PointPair
    {
        public PointPair(double x1, double y1, double x2, double y2)
            : this(new Point(x1, y1), new Point(x2, y2))
        {
        }

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
                result.P1 = result.P1.WithX(result.P2.X);
                result.P2 = result.P2.WithX(t);
            }
            if (result.P1.Y > result.P2.Y)
            {
                var t = result.P1.Y;
                result.P1 = result.P1.WithY(result.P2.Y);
                result.P2 = result.P2.WithY(t);
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

        public PointPair Inflate(double size)
        {
            return new PointPair(P1.Offset(-size), P2.Offset(size));
        }

        public bool HasValidValue()
        {
            return ((P1.X.IsValidValue() && P1.Y.IsValidValue()) || (P2.X.IsValidValue() && P2.Y.IsValidValue()));
        }
        public Point FirstValidValue()
        {
            if (P1.X.IsValidValue() && P1.Y.IsValidValue()) return P1;
            if (P2.X.IsValidValue() && P2.Y.IsValidValue()) return P2;
            return Math.InfinitePoint;
        }
    }

    public static partial class PointExtensions
    {
        public static double Distance(this Point p1, Point p2)
        {
            var x = p1.X - p2.X;
            x = x * x;
            var y = p1.Y - p2.Y;
            y = y * y;
            return System.Math.Sqrt(x + y);
        }

        public static double AngleTo(this Point center, Point point)
        {
            return Math.OAngle(center.Plus(new Point(10, 0)), center, point);
        }

        public static Point Reflect(this Point point, Point center)
        {
            return new Point(2 * center.X - point.X, 2 * center.Y - point.Y);
        }

        public static double Length(this Point p)
        {
            return p.SumOfSquares().SquareRoot();
        }

        public static double Arg(this Point p)
        {
            return new Point().AngleTo(p);
        }

        public static Point Scale(this Point p, double scaleFactor)
        {
            return new Point(p.X * scaleFactor, p.Y * scaleFactor);
        }

        public static Point SnapToIntegers(this Point p)
        {
            return new Point(M.Ceiling(p.X), M.Ceiling(p.Y));
        }

        public static Point PointInDirection(this Point start, Point direction, double vectorLength)
        {
            var vector = direction.Minus(start);
            var factor = vectorLength / vector.Length();
            return start.Plus(vector.Scale(factor));
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

        public static Point Minus(this Point point, double offset)
        {
            return new Point(point.X - offset, point.Y - offset);
        }

        public static Point Plus(this Point point, double offset)
        {
            return new Point(point.X + offset, point.Y + offset);
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

        public static Point Offset(this Point point, double offset)
        {
            return new Point(point.X + offset, point.Y + offset);
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
            var x = point.X;
            var y = point.Y;
            x = x * x;
            y = y * y;
            return x + y;
        }

        /// <summary>
        /// Finite coordinates. An infinite one (pi / (2 * F.X) with F.X = 0, the Sine Wave's
        /// crest at frequency 0) is as much "no point" as NaN: pushed to a canvas it made
        /// Avalonia throw "Invalid Arrange rectangle" and took the app down.
        /// </summary>
        public static bool Exists(this Point p)
        {
            return p.X.IsValidValue() && p.Y.IsValidValue();
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

        public static bool EqualsWithPrecision(this double value, double center)
        {
            return Math.Abs(value - center) < Math.Precision;
        }

        public static bool IsWithinEpsilonTo(this double value, double center)
        {
            return Math.Abs(value - center) < Math.Epsilon;
        }

        public static bool EqualsWithPrecision(this Point point, Point other)
        {
            return point.X.EqualsWithPrecision(other.X) && point.Y.EqualsWithPrecision(other.Y);
        }

        public static bool IsWithinEpsilon(this double value)
        {
            return Math.Abs(value) < Math.Epsilon;
        }

        public static bool IsWithinToleranceFactor(this double value, double factor)
        {
            return Math.Abs(value) < Math.CursorTolerance * factor;
        }

        public static bool IsWithinTolerance(this double value)
        {
            return IsWithinToleranceFactor(value, 1.0);
        }

        public static bool IsEqual(this Point point, Point other)
        {
            return point.Distance(other).IsWithinEpsilon();
        }

        // (noise off a coordinate, on every conversion from pixels: not the rounding for
        // display, which goes through decimal)
        public static double RoundToEpsilon(this double point)
        {
            return M.Round(point, 10);
        }

        public static Point RoundToEpsilon(this Point point)
        {
            return new Point(point.X.RoundToEpsilon(), point.Y.RoundToEpsilon());
        }

        public static IEnumerable<Point> RoundToEpsilon(this IEnumerable<Point> list)
        {
            return list.Select(p => p.RoundToEpsilon());
        }
    }

    public static partial class Math
    {
        public enum lengthUnit
        {
            Unitless = 0,
            Centimeter = 1,
            Inches = 2
        }

        public static double centimeterLogicalLength = .5906;

        public static double inchesLogicalLength = 1.5;

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

        /// <summary>
        /// http://www.ecse.rpi.edu/Homepages/wrf/Research/Short_Notes/pnpoly.html
        /// </summary>
        public static bool IsPointInPolygon(this IList<Point> polygon, Point start)
        {
            int n = polygon.Count;
            double x = start.X;
            double y = start.Y;
            bool inside = false;
            int i = 0;
            int j = 0;

            for (i = 0, j = n - 1; i < n; j = i++)
            {
                if (((polygon[i].Y > y) != (polygon[j].Y > y)) &&
                 (x < (polygon[j].X - polygon[i].X) * (y - polygon[i].Y) 
                    / (polygon[j].Y - polygon[i].Y) + polygon[i].X))
                {
                    inside = !inside;
                }
            }

            return inside;
        }

        public static bool IsPointInPolygonOld(this IList<Point> polygon, Point start)
        {
            if (polygon == null)
            {
                return false;
            }

            var outsidePoint = new Point(polygon.Max(p => p.X) + 100, 0);
            var ray = new PointPair(start, outsidePoint);

            int n = polygon.Count;
            int intersectionsCount = 0;
            for (int i = 0; i < n; i++)
            {
                PointPair side = new PointPair(polygon[i], polygon[i.RotateNext(n)]);

                Point intersectionPoint = GetIntersectionOfLines(ray, side);
                if (!intersectionPoint.Exists())
                {
                    continue;
                }

                intersectionPoint = intersectionPoint.RoundToEpsilon();

                bool isPointOnSide = IsPointInSegmentBoundingRect(side, intersectionPoint);
                bool isPointOnRay = IsPointInSegmentBoundingRect(ray, intersectionPoint);

                if (isPointOnSide
                    && isPointOnRay
                    && !intersectionPoint.IsEqual(side.P2))
                {
                    int angleWithPreviousSide = Math.VectorProduct(
                            start,
                            side.P1,
                            polygon[i.RotatePrevious(n)])
                        .Sign();
                    int angleWithNextSide = Math.VectorProduct(
                            start,
                            side.P1,
                            side.P2)
                        .Sign();
                    if (intersectionPoint.IsEqual(side.P1)
                        && angleWithPreviousSide == angleWithNextSide)
                    {
                        continue;
                    }

                    intersectionsCount++;
                }
            }

            return intersectionsCount % 2 == 1;
        }

        public static List<Point> GetIntersectionsOfPolygonAndSegment(IList<Point> polygon, PointPair segment, bool inclusive)
        {
            List<Point> result = new List<Point>();
            int n = polygon.Count;
            for (int i = 0; i < n; i++)
            {
                PointPair side = new PointPair(polygon[i], polygon[i.RotateNext(n)]);
                Point intersection = InfinitePoint;
                intersection = GetIntersectionOfSegments(side, segment, inclusive);
                if (intersection.Exists())
                {
                    if (inclusive)
                    {
                        result.Add(intersection);
                    }
                    else
                    {
                        if (!intersection.IsEqual(polygon[i.RotateNext(n)])
                            && !intersection.IsEqual(segment.P1)
                            && !intersection.IsEqual(segment.P2))
                        {
                            result.Add(intersection);
                        }
                    }
                }
            }
            return result;
        }

        public static List<Point> GetIntersectionsOfPolygonAndLine(IList<Point> polygon, PointPair line, bool inclusive)
        {
            List<Point> result = new List<Point>();
            int n = polygon.Count;
            for (int i = 0; i < n; i++)
            {
                PointPair side = new PointPair(polygon[i], polygon[i.RotateNext(n)]);
                Point intersection = GetIntersectionOfSegmentAndLine(side, line, inclusive);
                if (intersection.Exists())
                {
                    if (inclusive)
                    {
                        // When inclusive prevent redundant intersections by ignoring intersections at the side endpoint.
                        bool onSideEndPoint = Distance(intersection, side.P2).IsWithinEpsilon();
                        if (!onSideEndPoint)
                        {
                            result.Add(intersection);
                        }
                    }
                    else if (!intersection.IsEqual(polygon[i.RotateNext(n)])
                            && !intersection.IsEqual(line.P1)
                            && !intersection.IsEqual(line.P2))
                    {
                        result.Add(intersection);
                    }

                }
            }
            return result;
        }

        public static List<Point> GetIntersectionsOfPolygonAndPolygon(IList<Point> polygon1, IList<Point> polygon2, bool inclusive)
        {
            List<Point> result = new List<Point>();
            int n = polygon1.Count;
            for (int i = 0; i < n; i++)
            {
                PointPair side = new PointPair(polygon1[i], polygon1[i.RotateNext(n)]);
                List<Point> intersections = GetIntersectionsOfPolygonAndSegment(polygon2, side, inclusive);
                result.AddRange(intersections);
            }
            return result;
        }

        public static List<Point> GetIntersectionsOfPolygonAndEllipse(IList<Point> polygon, IEllipse ellipse, bool inclusive)
        {
            List<Point> result = new List<Point>();
            int n = polygon.Count;
            for (int i = 0; i < n; i++)
            {
                PointPair side = new PointPair(polygon[i], polygon[i.RotateNext(n)]);
                PointPair intersections = GetIntersectionOfEllipseAndSegment(ellipse, side);
                if (intersections.P1.Exists()) result.Add(intersections.P1);
                if (intersections.P2.Exists()) result.Add(intersections.P2);
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

        public static PointPair GetPerpendicularLine(PointPair parentLine, Point point)
        {
            PointPair coordinates = new PointPair()
            {
                P1 = point,
                P2 = new Point(
                    point.X + parentLine.P2.Y - parentLine.P1.Y,
                    point.Y + parentLine.P1.X - parentLine.P2.X)
            };
            return coordinates;
        }

        public static int Sign(this double num)
        {
            return M.Sign(num);
        }

        public static double Abs(this double num)
        {
            return M.Abs(num);
        }

        /// <summary>
        /// Rounded for showing. Plus zero: -0.001 rounds to a negative zero, which is
        /// written "-0" (a label said A is at (3, -0)); adding zero makes it a plain one.
        /// </summary>
        /// <remarks>
        /// A 5 rounds up, as at school: 3.125 is 3.13. (The default of System.Math is the
        /// banker's, to the even digit: an area of 3.125 was shown as 3.12, and of the
        /// numbers a point snapped to the grid gives, every other half went the wrong way.)
        /// Through decimal, which takes the number as it is written: 2.675 as a double is a
        /// hair under and would still round down.
        /// </remarks>
        public static double Round(this double num, int fractionalDigits)
        {
            if (M.Abs(num) < 1e15)
            {
                return (double)M.Round((decimal)num, fractionalDigits, System.MidpointRounding.AwayFromZero) + 0.0;
            }

            return M.Round(num, fractionalDigits) + 0.0;
        }

        public static double Round(this double num)
        {
            return M.Round(num);
        }

        public static double CursorTolerance
        {
            get
            {
                return Settings.Instance.CursorTolerance;
            }
            set
            {
                Settings.Instance.CursorTolerance = value;
            }
        }

        public static Point ScalePointBetweenTwo(Point p1, Point p2, double ratio)
        {
            return new Point(
                p1.X + (p2.X - p1.X) * ratio,
                p1.Y + (p2.Y - p1.Y) * ratio);
        }

        public static Point ScalePointBetweenTwo(PointPair segment, double ratio)
        {
            return ScalePointBetweenTwo(segment.P1, segment.P2, ratio);
        }

        public static double ToDegrees(this double radians)
        {
            return radians * 180 / PI;
        }

        public static string ToDegreeString(this double degrees)
        {
            return string.Format("{0:F0}°", degrees);
        }

        public static double ToRadians(this double degrees)
        {
            return degrees * PI / 180;
        }

        public static string ToRadianString(this double radians)
        {
            return string.Format("{0:F02} rad", radians);
        }

        public static double OAngle(Point firstPoint, Point vertex, Point secondPoint)
        {
            var a1 = GetAngle(vertex, firstPoint);
            var a2 = GetAngle(vertex, secondPoint);
            if (a2 < a1)
            {
                a2 = a2 + 2 * PI;
            }
            var result = a2 - a1;
            if (result > 2 * PI)
            {
                result -= 2 * PI;
            }

            // A full turn but for a rounding error is no turn: two sides along one ray (a
            // point put on ray BA) came out 0 or 359.99999 at random, and the angle's number
            // and mark blinked between 0 and 360 degrees while anything moved.
            if (2 * PI - result < 1e-9)
            {
                result = 0;
            }

            return result;
        }

        public static double GetAngle(double direction1, double direction2)
        {
            // Get the angle between the two directions (or angles).
            // This is analogous to finding the direction between two points.
            var angularSeparation = (direction2 % DOUBLEPI) - (direction1 % DOUBLEPI);
            return (angularSeparation.IsWithinEpsilon()) ? 0 : angularSeparation;
        }

        public static double GetAngle(Point center, Point endPoint)
        {
            var result = M.Atan2(endPoint.Y - center.Y, endPoint.X - center.X);
            if (result < 0)
            {
                result += 2 * PI;
            }
            return result;
        }

        public static Point GetOrthoPosition(Point center, Point point)
        {
            double angle = Math.GetAngle(center, point);
            if ((angle > Math.PI / 4 && angle < (3 * Math.PI / 4)) || (angle > (5 * Math.PI / 4) && angle < (7 * Math.PI / 4)))
            {
                point = point.WithX(center.X);
            }
            else
            {
                point = point.WithY(center.Y);
            }
            return point;
        }

        public static Point GetRotationPoint(Point p, Point center, double angle)
        {
            Point result = new Point(center.X, center.Y);

            if (angle == 0) return p;   // Rotate a point 0 degrees - no change.
            if (p == center) return p;  // Rotate a point about itself - no change.

            var r = Distance(center, p);
            var curAngle = M.Atan2(p.Y - center.Y, p.X - center.X);
            var newAngle = curAngle + angle;

            result += new Avalonia.Vector(r * M.Cos(newAngle), r * M.Sin(newAngle));

            return result;
        }

        public static Point GetSnapToGridPosition(double gridSpacing, Point point)
        {
            point = point.WithX(gridSpacing * M.Round(point.X / gridSpacing));
            point = point.WithY(gridSpacing * M.Round(point.Y / gridSpacing));

            // The values obtained with this code result in quirky behavior with non-integer values of gridSpacing.
            //if (point.X != 0)
            //    point.X = gridSpacing * (int)((point.X + (0.5 * point.X / M.Abs(point.X))) / gridSpacing);
            //else
            //    point.X = 0;

            //if (point.Y != 0)
            //    point.Y = gridSpacing * (int)((point.Y + (0.5 * point.Y / M.Abs(point.Y))) / gridSpacing);
            //else
            //    point.Y = 0;

            return point;
        }

        public static Point GetSnapToPointPosition(double gridSpacing, Point point, List<Point> points, bool withSnapToGrid)
        {
            foreach (Point pt in points)
            {
                if (point.Distance(pt) <= ((double)gridSpacing / 4))
                {
                    point = point.WithX(pt.X);
                    point = point.WithY(pt.Y);
                    return point;
                }
            }

            if (withSnapToGrid)
                point = GetSnapToGridPosition(gridSpacing, point);

            return point;
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

        public static double DOUBLEPI
        {
            get
            {
                return 2 * M.PI;
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
                result.P1 = result.P1.WithX(segment.P1.X);
                result.P2 = result.P2.WithX(segment.P1.X);
                result.P1 = result.P1.WithY(borders.P1.Y);
                result.P2 = result.P2.WithY(borders.P2.Y);
                if (segment.P1.Y > segment.P2.Y)
                {
                    result.P1 = result.P1.WithY(borders.P2.Y);
                    result.P2 = result.P2.WithY(borders.P1.Y);
                }
            }
            else if (segment.P1.Y == segment.P2.Y)
            {
                result.P1 = result.P1.WithX(borders.P1.X);
                result.P1 = result.P1.WithY(segment.P1.Y);
                result.P2 = result.P2.WithX(borders.P2.X);
                result.P2 = result.P2.WithY(segment.P1.Y);
                if (segment.P1.X > segment.P2.X)
                {
                    result.P1 = result.P1.WithX(borders.P2.X);
                    result.P2 = result.P2.WithX(borders.P1.X);
                }
            }
            else
            {
                var deltaX = segment.P2.X - segment.P1.X;
                var deltaY = segment.P2.Y - segment.P1.Y;
                var deltaXYRatio = deltaX / deltaY;
                var deltaYXRatio = deltaY / deltaX;

                result.P1 = result.P1.WithY(deltaY > 0 ? borders.P1.Y : borders.P2.Y);
                result.P1 = result.P1.WithX(segment.P1.X + (result.P1.Y - segment.P1.Y) * deltaXYRatio);
                if (result.P1.X < borders.P1.X)
                {
                    result.P1 = result.P1.WithX(borders.P1.X);
                    result.P1 = result.P1.WithY(segment.P1.Y + (result.P1.X - segment.P1.X) * deltaYXRatio);
                }
                else if (result.P1.X > borders.P2.X)
                {
                    result.P1 = result.P1.WithX(borders.P2.X);
                    result.P1 = result.P1.WithY(segment.P1.Y + (result.P1.X - segment.P1.X) * deltaYXRatio);
                }

                result.P2 = result.P2.WithX(deltaX > 0 ? borders.P2.X : borders.P1.X);
                result.P2 = result.P2.WithY(segment.P2.Y + (result.P2.X - segment.P2.X) * deltaYXRatio);
                if (result.P2.Y < borders.P1.Y)
                {
                    result.P2 = result.P2.WithY(borders.P1.Y);
                    result.P2 = result.P2.WithX(segment.P2.X + (result.P2.Y - segment.P2.Y) * deltaXYRatio);
                }
                else if (result.P2.Y > borders.P2.Y)
                {
                    result.P2 = result.P2.WithY(borders.P2.Y);
                    result.P2 = result.P2.WithX(segment.P2.X + (result.P2.Y - segment.P2.Y) * deltaXYRatio);
                }
            }

            return result;
        }

        public struct ProjectionInfo
        {
            public Point Point { get; set; }
            public double Ratio { get; set; }
            public double DistanceToLine { get; set; }

            public bool IsWithinSegment()
            {
                return Ratio >= 0 && Ratio <= 1;
            }
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

        static double GetProjectionRatio(PointPair line, Point projection)
        {
            var result = 0.0;
            // Need to use epsilon to prevent erroneous result.
            if (!(line.P1.X - line.P2.X).IsWithinEpsilon())
            {
                result = (projection.X - line.P1.X) / (line.P2.X - line.P1.X);
            }
            else if (!(line.P1.Y - line.P2.Y).IsWithinEpsilon())
            {
                result = (projection.Y - line.P1.Y) / (line.P2.Y - line.P1.Y);
            }
            return result;
        }

        /// <summary>
        /// How far the point is from the line, measured along the perpendicular (from P1
        /// when the line has no direction).
        /// </summary>
        public static double GetDistanceToLine(Point point, PointPair line)
        {
            return Distance(point, GetProjectionPoint(point, line));
        }

        public static Point GetProjectionPoint(Point p, PointPair line)
        {
            Point result = new Point();

            if (line.P1.Y.IsWithinEpsilonTo(line.P2.Y))
            {
                if (line.P1.X.IsWithinEpsilonTo(line.P2.X))
                {
                    result = line.P1;
                }
                else
                {
                    result = result.WithX(p.X);
                    result = result.WithY(line.P1.Y);
                }
            }
            else if (line.P1.X.IsWithinEpsilonTo(line.P2.X))
            {
                result = result.WithX(line.P1.X);
                result = result.WithY(p.Y);
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

        public static ProjectionInfo GetProjection(Point point, IList<Point> polygonalChain, bool IsClosed)
        {
            ProjectionInfo nearestProjection = new ProjectionInfo();
            nearestProjection.DistanceToLine = double.MaxValue;
            var count = polygonalChain.Count;
            var stop = (IsClosed) ? count : count - 1;
            for (int i = 0; i < stop; i++)
            {
                var segment = new PointPair(polygonalChain[i], polygonalChain[i.RotateNext(count)]);
                var projectionInfo = GetProjection(point, segment);
                if ((projectionInfo.DistanceToLine < nearestProjection.DistanceToLine)
                        && projectionInfo.IsWithinSegment())
                {
                    nearestProjection = projectionInfo;
                }
            }

            return nearestProjection;
        }

        /// <summary>
        /// Get the point in pointList nearest to given point.
        /// </summary>
        public static Point GetNearestPoint(Point point, IList<Point> pointList)
        {
            return pointList[GetIndexOfNearestPoint(point, pointList)];
        }

        /// <summary>
        /// Get the index of the point in pointList nearest to given point.
        /// </summary>
        public static int GetIndexOfNearestPoint(Point point, IList<Point> pointList)
        {
            double nearest = double.PositiveInfinity;
            Point nearestPoint = new Point();
            int nearestPointIndex = 0;
            for (int i = 0; i < pointList.Count(); i++)
            {
                var currentPoint = pointList[i];
                var currentDistance = point.Distance(currentPoint);
                if (currentDistance < nearest)
                {
                    nearest = currentDistance;
                    nearestPoint = currentPoint;
                    nearestPointIndex = i;
                }
            }

            return nearestPointIndex;
        }

        public struct BezierInfo
        {
            double ax;
            double ay;
            double bx;
            double by;
            double cx;
            double cy;

            Point p0;
            Point p1;
            Point p2;
            Point p3;

            public readonly Point[] Points;

            public const int NumberOfPoints = 50;

            public BezierInfo(
                Point p0,
                Point p1,
                Point p2,
                Point p3)
            {
                this.p0 = p0;
                this.p1 = p1;
                this.p2 = p2;
                this.p3 = p3;

                cx = 3 * (p1.X - p0.X);
                bx = 3 * (p2.X - p1.X) - cx;
                ax = p3.X - p0.X - cx - bx;

                cy = 3 * (p1.Y - p0.Y);
                by = 3 * (p2.Y - p1.Y) - cy;
                ay = p3.Y - p0.Y - cy - by;

                Points = null;
                Points = GetPoints();
            }

            public Point GetPoint(double t)
            {
                Point result = new Point();
                double t2 = t * t;
                double t3 = t2 * t;
                result = result.WithX(ax * t3 + bx * t2 + cx * t + p0.X);
                result = result.WithY(ay * t3 + by * t2 + cy * t + p0.Y);
                return result;
            }

            Point[] GetPoints()
            {
                Point[] result = new Point[NumberOfPoints];
                double precisionMinus1 = NumberOfPoints - 1;
                for (int i = 0; i < NumberOfPoints; i++)
                {
                    double t = i / precisionMinus1;
                    result[i] = GetPoint(t);
                }

                return result;
            }

            public ProjectionInfo GetProjection(Point point)
            {
                return Math.GetProjection(point, Points, false);
            }

            public double GetNearestParameterFromPoint(Point point)
            {
                return Math.GetNearestParameterFromPointOnPolyline(Points, point);
            }
        }

        public static double GetNearestParameterFromPointOnPolyline(IList<Point> points, Point point)
        {
            ProjectionInfo nearestProjection = new ProjectionInfo();
            nearestProjection.DistanceToLine = double.MaxValue;
            int nearestIndex = 0;
            int nearestPointIndex = 0;
            int numberOfPoints = points.Count;
            double totalLength = 0;
            double parameter = 0;
            double vertexParameter = 0;
            double nearestDistance = double.PositiveInfinity;

            for (int i = 0; i < points.Count - 1; i++)
            {
                var segment = new PointPair(points[i], points[i + 1]);
                var projectionInfo = Math.GetProjection(point, segment);
                if ((projectionInfo.DistanceToLine < nearestProjection.DistanceToLine)
                    && projectionInfo.IsWithinSegment())
                {
                    nearestProjection = projectionInfo;
                    nearestIndex = i;
                    parameter = totalLength + segment.Length * projectionInfo.Ratio;
                }
                var distance = point.Distance(points[i]);
                if (distance < nearestDistance)
                {
                    nearestDistance = distance;
                    nearestPointIndex = i;
                    vertexParameter = totalLength;
                }
                totalLength += segment.Length;
            }

            if (point.Distance(points.Last()) < nearestDistance)
            {
                nearestDistance = point.Distance(points.Last());
                nearestPointIndex = points.Count - 1;
                vertexParameter = totalLength;
            }

            if (nearestDistance < nearestProjection.DistanceToLine)
            {
                return vertexParameter / totalLength;
            }

            return parameter / totalLength;
        }

        public static bool IsPointOnLine(PointPair line, Point point, double epsilon)
        {
            var projection = GetProjection(point, line);
            return projection.DistanceToLine < epsilon;
        }

        public static bool ArePointsCollinear(Point a, Point b, Point c)
        {
            if (a.EqualsWithPrecision(b) 
                || b.EqualsWithPrecision(c)
                || c.EqualsWithPrecision(a))
            {
                return false;
            }

            return IsPointOnLine(new PointPair(a, b), c, Math.Epsilon);
        }

        public static bool IsPointOnSegment(PointPair line, Point point, double epsilon)
        {
            var projection = GetProjection(point, line);
            var slack = EndTolerance(line);
            return projection.DistanceToLine < epsilon
                && projection.Ratio >= -slack
                && projection.Ratio <= 1 + slack;
        }

        /// <summary>
        /// Is the given point on the polygonal chain? IsClosed = false for polylines and beziers, true for polygons.
        /// </summary>
        public static bool IsPointOnPolygonalChain(IList<Point> points, Point point, double epsilon, bool IsClosed)
        {
            // a curve not worked out yet (a function graph while its file is being read)
            if (points == null || points.Count == 0)
            {
                return false;
            }

            var projection = Math.GetProjection(point, points, IsClosed);
            var projectionDistance = projection.DistanceToLine;
            if (projectionDistance < epsilon)
            {
                return true;
            }

            var nearestPoint = Math.GetNearestPoint(point, points);
            var nearestPointDistance = nearestPoint.Distance(point);
            if (nearestPointDistance < epsilon)
            {
                return true;
            }

            return false;
        }

        public static Point GetAngleBisectorPoint(Point vertex, Point side1, Point side2)
        {
            var s1 = vertex.Distance(side1);
            var s2 = vertex.Distance(side2);
            if (s1 == 0 || s2 == 0)
            {
                return Math.InfinitePoint;
            }

            var a1 = vertex.AngleTo(side1);
            var a2 = vertex.AngleTo(side2);
            if (a2 < a1)
            {
                a2 += 2 * PI;
            }
            var a = (a1 + a2) / 2;

            return Math.RotatePoint(vertex, vertex.Distance(side1), a);
        }

        public static Point GetIntersectionOfSegments(PointPair segment1, PointPair segment2, bool inclusive)
        {
            Point result = GetIntersectionOfLines(segment1, segment2);
            if (!result.Exists())
            {
                return result;
            }
            if (inclusive)
            {
                if (IsPointInSegmentBoundingRect(segment1, result)
                    && IsPointInSegmentBoundingRect(segment2, result))
                {
                    return result;
                }
            }
            else
            {
                if (IsPointInSegmentInnerBoundingRect(segment1, result)
                    && IsPointInSegmentInnerBoundingRect(segment2, result))
                {
                    return result;
                }
            }
            return Math.InfinitePoint;
        }

        public static Point GetIntersectionOfSegmentAndLine(PointPair segment, PointPair line, bool inclusive)
        {
            Point result = GetIntersectionOfLines(segment, line);
            if (!result.Exists())
            {
                return result;
            }
            if (inclusive)
            {
                if (IsPointInSegmentBoundingRect(segment, result))
                {
                    return result;
                }
            }
            else
            {
                if (IsPointInSegmentInnerBoundingRect(segment, result))
                {
                    return result;
                }
            }
            return Math.InfinitePoint;
        }

        public static PointPair GetIntersectionOfSegmentAndRect(PointPair segment, Avalonia.Rect rect)
        {
            PointPair result = GetIntersectionOfLineAndRect(segment, rect);
            if (!result.HasValidValue())
            {
                return result;
            }
            if (IsPointInSegmentInnerBoundingRect(segment, result.P1) || IsPointInSegmentInnerBoundingRect(segment, result.P2))
            {
                return result;
            }
            return Math.InfinitePointPair;
        }

        public static PointPair GetIntersectionOfLineAndRect(PointPair line, Avalonia.Rect rect)
        {
            // Using logical coordinate system.
            var intersections = InfinitePointPair;
            Point TL = new Point(rect.X, rect.Y);
            Point TR = new Point(rect.X + rect.Width, rect.Y);
            Point BR = new Point(rect.X + rect.Width, rect.Y - rect.Height);
            Point BL = new Point(rect.X, rect.Y - rect.Height);
            Point[] ints = new Point[4];
            ints[0] = GetIntersectionOfSegmentAndLine(new PointPair(TL, TR), line, true);
            ints[1] = GetIntersectionOfSegmentAndLine(new PointPair(TR, BR), line, true);
            ints[2] = GetIntersectionOfSegmentAndLine(new PointPair(BR, BL), line, true);
            ints[3] = GetIntersectionOfSegmentAndLine(new PointPair(BL, TL), line, true);
            for (int i = 0; i < 4; i++)
            {
                if (ints[i].X.IsValidValue() && ints[i].Y.IsValidValue())
                {
                    if (!intersections.P1.X.IsValidValue() && !intersections.P1.Y.IsValidValue())
                    {
                        intersections.P1 = ints[i];
                    }
                    else
                    {
                        intersections.P2 = ints[i];
                    }
                }
            }
            return intersections;
        }

        public static bool IsPointInSegmentBoundingRect(PointPair segment, Point point)
        {
            var boundingRect = segment.GetBoundingRect().Inflate(Math.Epsilon);
            return boundingRect.Contains(point);
        }

        public static bool IsPointInSegmentInnerBoundingRect(PointPair segment, Point point)
        {
            segment = segment.GetBoundingRect();
            if (segment.P1.X.IsWithinEpsilonTo(segment.P2.X))
            {
                return point.X.EqualsWithPrecision(segment.P1.X)
                    && point.Y > segment.P1.Y
                    && point.Y < segment.P2.Y;
            }
            else if (segment.P1.Y.IsWithinEpsilonTo(segment.P2.Y))
            {
                return point.Y.EqualsWithPrecision(segment.P1.Y)
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

            // parallel by the sine of the angle between the lines, not by the determinant, which
            // grows with their lengths: two short lines at a right angle passed for parallel
            var lengths = M.Sqrt(a1 * a1 + b1 * b1) * M.Sqrt(a2 * a2 + b2 * b2);
            if (M.Abs(d) <= Epsilon * lengths)
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
            Point[] array = points.ToArray();
            if (array == null || array.Length == 0)
            {
                return new Point();
            }
            Point sum = new Point();
            foreach (var point in array)
            {
                sum = sum.Plus(point);
            }
            return sum.Scale(1.0 / array.Length);
        }

        public static Point Midpoint(this IEnumerable<IPoint> points)
        {
            return points.Select(p => p.Coordinates).Midpoint();
        }

        public static Point Midpoint(params Point[] points)
        {
            return points.Midpoint();
        }

        //public static double Area(this IEnumerable<Point> vertices, CoordinateSystem reference)
        //{
        //    return reference.ToLogical(vertices).Area();
        //}

        public static double Area(params Point[] points)
        {
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
                // A polygon whose sides cross (a vertex dragged over the opposite side: a
                // bow tie) is measured as it is filled. The formula below takes the parts
                // that go round the other way away from the rest: the two halves of a
                // symmetric bow tie came to 0, under a figure plainly filled.
                var levels = CrossingLevels(points);
                if (levels != null)
                {
                    return FilledArea(points, levels);
                }

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

        /// <summary>
        /// The heights at which two sides of the polygon cross each other (not at a vertex
        /// they share); null for a polygon whose sides don't cross
        /// </summary>
        static List<double> CrossingLevels(Point[] points)
        {
            List<double> levels = null;
            int count = points.Length;
            for (int i = 0; i < count; i++)
            {
                var p = points[i];
                var r = points[(i + 1) % count].Minus(p);
                for (int j = i + 2; j < count; j++)
                {
                    // the last side and the first share a vertex too
                    if (i == 0 && j == count - 1)
                    {
                        continue;
                    }

                    var q = points[j];
                    var s = points[(j + 1) % count].Minus(q);
                    double denominator = r.X * s.Y - r.Y * s.X;
                    if (denominator == 0)
                    {
                        continue;
                    }

                    double t = ((q.X - p.X) * s.Y - (q.Y - p.Y) * s.X) / denominator;
                    double u = ((q.X - p.X) * r.Y - (q.Y - p.Y) * r.X) / denominator;
                    if (t > 0 && t < 1 && u > 0 && u < 1)
                    {
                        levels = levels ?? new List<double>();
                        levels.Add(p.Y + t * r.Y);
                    }
                }
            }

            return levels;
        }

        /// <summary>
        /// The area an even-odd fill of the polygon covers (how a polygon is filled on
        /// screen), in slabs between the heights of its vertices and of the crossings of
        /// its sides: within a slab no two sides cross, so what is filled there is a row of
        /// trapezoids, as wide in all as they are half way up.
        /// </summary>
        static double FilledArea(Point[] points, List<double> levels)
        {
            foreach (var point in points)
            {
                levels.Add(point.Y);
            }

            levels.Sort();
            double area = 0;
            var crossings = new List<double>();
            for (int level = 1; level < levels.Count; level++)
            {
                double height = levels[level] - levels[level - 1];
                if (!(height > 0))
                {
                    continue;
                }

                double middle = (levels[level] + levels[level - 1]) / 2;
                crossings.Clear();
                for (int i = 0; i < points.Length; i++)
                {
                    var a = points[i];
                    var b = points[(i + 1) % points.Length];
                    if ((a.Y < middle) != (b.Y < middle))
                    {
                        crossings.Add(a.X + (b.X - a.X) * (middle - a.Y) / (b.Y - a.Y));
                    }
                }

                crossings.Sort();
                for (int i = 0; i + 1 < crossings.Count; i += 2)
                {
                    area += (crossings[i + 1] - crossings[i]) * height;
                }
            }

            return area;
        }

        public static double Area(this IEnumerable<Point> vertices)
        {
            return Area(vertices.ToArray());
        }

        public static double Length(params Point[] points)
        {
            if (points.Length < 2)
            {
                return 0;
            }
            else
            {
                double sum = 0;
                for (int i = 0; i < points.Length - 1; i++)
                {
                    sum += points[i].Distance(points[i + 1]);
                }
                return sum;
            }
        }

        public static double Length(this IEnumerable<Point> polyline)
        {
            return Length(polyline.ToArray());
        }

        public const double Epsilon = 0.00000001;

        public const double Precision = 0.00000001;

        public static Point GetDilationPoint(Point p, Point center, double factor)
        {
            Point result = new Point(center.X, center.Y);

            if (factor == 0) return result; // Dilation point becomes center.

            var beforeDistance = Distance(center, p);

            if (beforeDistance == 0) return result; // Avoids divide by zero.

            var afterDistance = beforeDistance * factor;
            var dx = p.X - center.X;
            var dy = p.Y - center.Y;

            result += new Avalonia.Vector(afterDistance * dx / beforeDistance, afterDistance * dy / beforeDistance);

            return result;
        }

        /// <summary>
        /// How far apart two lengths near <paramref name="scale"/> may be and still count as
        /// equal when deciding whether figures touch. A line tangent to a circle by construction
        /// (or two circles) comes out a hair apart or a hair overlapping, at random as the
        /// figures move, and without this the touching point blinks in and out of existence.
        /// Relative, because the rounding error grows with the size of the numbers.
        /// </summary>
        public static double TangencyTolerance(double scale)
        {
            return 1e-9 * M.Max(1, scale);
        }

        static double Magnitude(Point point)
        {
            return M.Max(M.Abs(point.X), M.Abs(point.Y));
        }

        /// <summary>
        /// How far a point is from an ellipse, along the ray from the center through the
        /// point: negative inside, the radius minus the distance for a circle. What a hit
        /// test compares with the reach of the cursor. (It compared the left side of the
        /// ellipse's equation minus 1 instead, a number without units against a length: the
        /// band around a circle that took a click was as many times too wide as half the
        /// radius measures in units - a circle of radius 100 took every click on the screen,
        /// one of radius 0.2 almost none.)
        /// </summary>
        public static double RadialDistanceToEllipse(
            Point center,
            double semiMajor,
            double semiMinor,
            double inclination,
            Point point)
        {
            var distance = Distance(center, point);
            if (distance == 0)
            {
                return -M.Min(semiMajor, semiMinor);
            }

            // (distance from the center / distance of the edge in that direction) squared
            var angle = GetAngle(center, point) - inclination;
            var x = distance * M.Cos(angle);
            var y = distance * M.Sin(angle);
            var equationLeft = x * x / (semiMajor * semiMajor) + y * y / (semiMinor * semiMinor);
            return distance - distance / M.Sqrt(equationLeft);
        }

        /// <summary>
        /// How far past its ends a point may be and still be on a segment (or before the start
        /// of a ray), as a fraction of its length. An intersection that falls exactly on an end
        /// (a circle through the end of the segment) comes out a hair inside or outside, and
        /// without this it blinks in and out as the figures move; see <see cref="TangencyTolerance"/>.
        /// </summary>
        public static double EndTolerance(PointPair line)
        {
            return TangencyTolerance(M.Max(Magnitude(line.P1), Magnitude(line.P2))) / line.Length;
        }

        /// <returns>Intersections - inbound first, outbound second based on the order of points in line.</returns>
        public static PointPair GetIntersectionOfCircleAndLine(
            Point center,
            double radius,
            PointPair line)
        {
            var result = Math.InfinitePointPair;
            var dx = line.P2.X - line.P1.X;
            var dy = line.P2.Y - line.P1.Y;
            var lengthSquared = dx * dx + dy * dy;
            if (lengthSquared == 0)
            {
                return result;
            }

            // the foot of the perpendicular from the center, and the distance to it
            var t = ((center.X - line.P1.X) * dx + (center.Y - line.P1.Y) * dy) / lengthSquared;
            var p = new Point(line.P1.X + dx * t, line.P1.Y + dy * t);
            var h = center.Distance(p);

            var tolerance = TangencyTolerance(M.Max(radius, Magnitude(center)));
            if (h > radius + tolerance)
            {
                return result;
            }

            // touching: the foot itself, exactly - no square root of a rounding error, which
            // would shake whatever is built on the point
            double s = 0;
            if (h < radius - tolerance)
            {
                s = ((radius - h) * (radius + h)).SquareRoot() / lengthSquared.SquareRoot();
            }

            result.P1 = new Point(p.X - dx * s, p.Y - dy * s);
            result.P2 = new Point(p.X + dx * s, p.Y + dy * s);
            return result;
        }

        public static PointPair GetIntersectionOfCircleAndSegment(
            Point center,
            double radius,
            PointPair segment)
        {
            var result = InfinitePointPair;
            var ints = GetIntersectionOfCircleAndLine(center, radius, segment);
            if (IsPointInSegmentInnerBoundingRect(segment, ints.P1))
            {
                result.P1 = ints.P1;
            }
            if (IsPointInSegmentInnerBoundingRect(segment, ints.P2))
            {
                result.P2 = ints.P2;
            }
            return result;
        }

        public static PointPair GetIntersectionOfEllipseAndSegment(IEllipse ellipse, PointPair segment)
        {
            var result = InfinitePointPair;
            var ints = GetIntersectionOfEllipseAndLine(ellipse, segment);
            if (IsPointInSegmentInnerBoundingRect(segment, ints.P1))
            {
                result.P1 = ints.P1;
            }
            if (IsPointInSegmentInnerBoundingRect(segment, ints.P2))
            {
                result.P2 = ints.P2;
            }
            return result;
        }

        public static PointPair GetIntersectionOfEllipseAndSegment(
            Point center,
            double semiMajor,
            double semiMinor,
            double angle,
            PointPair segment)
        {
            var result = InfinitePointPair;
            var ints = GetIntersectionOfEllipseAndLine(center, semiMajor, semiMinor, angle, segment);
            if (IsPointInSegmentInnerBoundingRect(segment, ints.P1))
            {
                result.P1 = ints.P1;
            }
            if (IsPointInSegmentInnerBoundingRect(segment, ints.P2))
            {
                result.P2 = ints.P2;
            }
            return result;
        }

        public static PointPair GetIntersectionOfEllipseAndLine(IEllipse ellipse, PointPair line)
        {
            Point center = ellipse.Center;
            double semiMajor = ellipse.SemiMajor;
            double semiMinor = ellipse.SemiMinor;
            double angle = ellipse.Inclination;
            return GetIntersectionOfEllipseAndLine(center, semiMajor, semiMinor, angle, line);
        }

        public static PointPair GetIntersectionOfEllipseAndLine(
            Point center,
            double semiMajor,
            double semiMinor,
            double angle,
            double slope,
            Point point)
        {
            Point point2 = new Point(point.X + 1, point.Y + slope);
            return GetIntersectionOfEllipseAndLine(center, semiMajor, semiMinor, angle, new PointPair(point, point2));
        }

        public static PointPair GetIntersectionOfEllipseAndLine(
            Point center,
            double semiMajor,
            double semiMinor,
            double angle,
            PointPair line)
        {
            if (semiMajor == 0 || semiMinor == 0) return InfinitePointPair;
            var hScale = semiMajor / semiMinor;
            var hScaleInv = semiMinor / semiMajor;

            // Transform the line-ellipse system, treat as circle and line, then transform back.
            var p1 = RotatePoint(line.P1, center, -angle);
            var p2 = RotatePoint(line.P2, center, -angle);
            p1 = p1.Minus(center);
            p2 = p2.Minus(center);
            p1 = p1.WithX(p1.X * hScaleInv);
            p2 = p2.WithX(p2.X * hScaleInv);
            var transformedLine = new PointPair(p1, p2);

            var ints = GetIntersectionOfCircleAndLine(new Point(0, 0), semiMinor, transformedLine);

            ints.P1 = ints.P1.WithX(ints.P1.X * hScale);
            ints.P2 = ints.P2.WithX(ints.P2.X * hScale);
            ints.P1 = ints.P1.Plus(center);
            ints.P2 = ints.P2.Plus(center);
            ints.P1 = RotatePoint(ints.P1, center, angle);
            ints.P2 = RotatePoint(ints.P2, center, angle);
            return ints;

        }

        public static double Slope(Point p2, Point p1)
        {
            double result = double.NaN;
            if (!(p2.X - p1.X).IsWithinEpsilon())
            {
                result = ((p2.Y - p1.Y) / (p2.X - p1.X));
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
            var dx = center2.X - center1.X;
            var dy = center2.Y - center1.Y;
            var distance = center1.Distance(center2);
            if (distance == 0)
            {
                return result;
            }

            var tolerance = TangencyTolerance(M.Max(M.Max(radius1, radius2), M.Max(Magnitude(center1), Magnitude(center2))));
            var outerTouch = distance - (radius1 + radius2);
            var innerTouch = distance - M.Abs(radius1 - radius2);
            if (outerTouch > tolerance || innerTouch < -tolerance)
            {
                return result;
            }

            // along the line of centers from center1 to the chord, and half the chord; a touch
            // is exact, as with a line (GetIntersectionOfCircleAndLine)
            var along = ((radius1 - radius2) * (radius1 + radius2) + distance * distance) / (2 * distance);
            double halfChord = 0;
            if (outerTouch < -tolerance && innerTouch > tolerance)
            {
                halfChord = ((radius1 - along) * (radius1 + along)).SquareRoot();
                if (double.IsNaN(halfChord))
                {
                    halfChord = 0;
                }
            }
            else if (outerTouch >= -tolerance)
            {
                // touching from outside: radius1 along the way to center2
                along = radius1;
            }
            else
            {
                // one inside the other: the touching point is on the far side of the smaller
                along = radius1 >= radius2 ? radius1 : -radius1;
            }

            var ux = dx / distance;
            var uy = dy / distance;
            var foot = new Point(center1.X + ux * along, center1.Y + uy * along);

            // P1 to the right of the way from center1 to center2, P2 to the left
            result.P1 = new Point(foot.X + uy * halfChord, foot.Y - ux * halfChord);
            result.P2 = new Point(foot.X - uy * halfChord, foot.Y + ux * halfChord);
            return result;
        }

        public static double Distance(double x0, double y0, double x1, double y1)
        {
            var x = x1 - x0;
            x = x * x;
            var y = y1 - y0;
            y = y * y;
            return System.Math.Sqrt(x + y);
        }

        public static double Distance(Point p1, Point p2)
        {
            return p1.Distance(p2);
        }

        public static double Distance(PointBase p1, PointBase p2)
        {
            return p1.Coordinates.Distance(p2.Coordinates);
        }

        public static double Distance(this IEnumerable<Point> points)
        {
            int count = points.Count();
            Point[] array = points.ToArray();
            if (array == null || count == 0)
            {
                return 0;
            }
            double distance = 0;
            for (int i = 0; i < count; i++)
            {
                distance += Distance(array[i], array[i.RotateNext(count)]);
            }
            return distance;
        }

        public static PointPair GetTangentPoints(Point outside, Point center, double radius)
        {
            var distance = outside.Distance(center);
            if (distance == 0 || distance < radius)
            {
                return Math.InfinitePointPair;
            }
            var angle = System.Math.Acos(radius / distance);
            var originalAngle = System.Math.Atan2(outside.Y - center.Y, outside.X - center.X);
            var point1 = RotatePoint(center, radius, originalAngle + angle);
            var point2 = RotatePoint(center, radius, originalAngle - angle);
            return new PointPair(point1, point2);
        }

        public static Point GetTranslationPoint(Point p, double distance, double direction)
        {
            Point result = new Point(p.X, p.Y);

            if (distance == 0) return p;   // Rotate a point 0 degrees - no change.

            var dx = distance * M.Cos(direction);
            var dy = distance * M.Sin(direction);
            result += new Avalonia.Vector(dx, dy);

            return result;
        }

        public static Point RotatePoint(Point point, Point center, double angle)
        {
            // Rotate point about center by angle.
            var p = point.Minus(center);
            var q = new Point(
                p.X * M.Cos(angle) - p.Y * M.Sin(angle),
                p.X * M.Sin(angle) + p.Y * M.Cos(angle)
                );
            return q.Plus(center);
        }

        public static Point RotatePoint(Point center, double radius, double angle)
        {
            // Rotate point from 0 radians to given angle in radians along circle of given radius and center.
            return new Point(center.X + radius * M.Cos(angle),
                center.Y + radius * M.Sin(angle));
        }

        public static bool IsAngleBetweenAngles(double a, double a1, double a2, bool clockwise)
        {
            // at an end, also across 0 = 2pi: a point a hair below an arc that starts at 0
            // comes out just under 2pi
            if (M.Abs(M.IEEERemainder(a - a1, 2 * PI)) < Epsilon) return true;
            if (M.Abs(M.IEEERemainder(a - a2, 2 * PI)) < Epsilon) return true;
            if ((a2 - a1).IsWithinEpsilon()) return false;
            if (clockwise)
            {
                var temp = a1;
                a1 = a2;
                a2 = temp;
            }
            if (a2 > a1)
            {
                return a >= a1 && a <= a2;
            }
            else
            {
                if (a <= a2)
                {
                    return true;
                }
                return a >= a1;
            }
        }

        public static Point GetSymmetricPoint(Point source, Point mirror)
        {
            return new Point(
                2 * mirror.X - source.X,
                2 * mirror.Y - source.Y);
        }

        public static Point GetSymmetricPoint(Point source, PointPair mirror)
        {
            var projection = GetProjectionPoint(source, mirror);
            return GetSymmetricPoint(source, projection);
        }

        public static Point GetSymmetricPoint(Point source, Point center, double radius)
        {
            if (radius < Math.Epsilon)
            {
                return Math.InfinitePoint;
            }

            var centerToDistance = source.Distance(center);
            var newRadius = radius * radius / centerToDistance;
            return center.PointInDirection(source, newRadius);
        }

        public static Point GetPointOnPolylineFromParameter(IList<Point> logicalPoints, double parameter)
        {
            // a curve not worked out yet (a locus while its file is being read): no such point
            if (logicalPoints == null || logicalPoints.Count == 0)
            {
                return InfinitePoint;
            }

            double sum = 0;
            double totalLength = logicalPoints.PolylineLength();
            for (int i = 0; i < logicalPoints.Count - 1; i++)
            {
                var segment = new PointPair(logicalPoints[i], logicalPoints[i + 1]);
                var oldParameter = sum / totalLength;
                sum += segment.Length;
                var newParameter = sum / totalLength;
                if (newParameter > parameter)
                {
                    var lambda = (parameter - oldParameter) / (newParameter - oldParameter);
                    return Math.ScalePointBetweenTwo(segment, lambda);
                }
            }
            return logicalPoints[logicalPoints.Count - 1];
        }

        public static readonly Size InfiniteSize = new Size(double.PositiveInfinity, double.PositiveInfinity);

        // double value validation
        public static bool IsDoubleValid(string inputNumber)
        {
            bool isNumberValid = false;
            double number = -1.0;
            if (System.Double.TryParse(inputNumber, out number))
            {
                isNumberValid = true;
            }
            return isNumberValid;
        }

        /// <summary>
        /// Angle from horizontal.
        /// </summary>
        public static double OHAngle(Point vertex, Point secondPoint)
        {
            // create horizontal first line 
            Point firstPoint = new Point();
            firstPoint = firstPoint.WithX(vertex.X + 10.0);
            firstPoint = firstPoint.WithY(vertex.Y);

            var result = -(OAngle(secondPoint, vertex, firstPoint) - Math.ToRadians(180));
            return result;
        }

        // Exact Angle by User input
        public static Point GetPositionByExactAngle(Point center, Point point, double Angle)
        {
            Point ret;

            if (Angle != 90 && Angle != 270)
            {
                double dx = point.X - center.X;
                double dy = point.Y - center.Y;
                point = point.WithY(center.Y + dx * M.Tan(Math.ToRadians(Angle)));
            }
            else
            {
                point = GetOrthoPosition(center, point);
            }

            ret = point;
            return ret;
        }

        // Exact Angle and Exact Length
        public static Point GetPositionByExactAngleAndLength(Point center, Point point, double Angle, double Length)
        {
            Point ret;

            double dx = Length * M.Cos(Math.ToRadians(Angle));
            double dy = Length * M.Sin(Math.ToRadians(Angle));

            point = point.WithX(center.X + dx);
            point = point.WithY(center.Y + dy);

            ret = point;
            return ret;
        }

        // For Snap to Center Line
        public static Point GetSnapToSegmentCenterPosition(Point point, List<Segment> segments)
        {
            foreach (Segment sg in segments)
            {
                if (sg != null)
                {
                    double tolerance = 0.1;
                    double P1x = sg.Coordinates.P1.X;
                    double P1y = sg.Coordinates.P1.Y;
                    double P2x = sg.Coordinates.P2.X;
                    double P2y = sg.Coordinates.P2.Y;
                    double Px = point.X;
                    double Py = point.Y;

                    double deltaX1 = M.Abs(P2x - Px);
                    double deltaY1 = M.Abs(P2y - Py);

                    double deltaX2 = M.Abs(P1x - Px);
                    double deltaY2 = M.Abs(P1y - Py);


                    if (deltaX1 > tolerance && deltaY1 > tolerance && deltaX2 > tolerance && deltaY2 > tolerance)
                    {
                        //search segment center
                        point = Math.ScalePointBetweenTwo(sg.Coordinates, 0.5);
                        return point;
                    }
                }
            }

            return point;
        }

        public static bool IsPointOnPolyline(IList<Point> points, Point point, double epsilon)
        {
            return IsPointOnPolygonalChain(points, point, epsilon, false);
        }
    }
}
