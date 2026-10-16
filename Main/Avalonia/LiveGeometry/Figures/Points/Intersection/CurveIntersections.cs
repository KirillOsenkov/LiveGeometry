using System;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// Where a curve crosses another figure: a function graph or a Bézier curve or path, walked
/// along, against a line, ray, segment, vector or axis, a circle, arc or ellipse, or a graph
/// (a graph with a graph, a Bézier curve with a graph; not two Bézier curves). Such figures
/// can cross any number of times, so the point keeps where on the curve it is
/// (<see cref="IntersectionPoint.Parameter"/>: x on a graph, t on a Bézier curve, the piece
/// and t on a path) and takes the crossing nearest that each time it is worked out: it
/// follows its own crossing as the figures move. A crossing is a root of the other figure's
/// equation along the curve (<see cref="Equation"/>, a signed distance or near it), looked
/// for outwards from the parameter on both sides at once and then narrowed down by halving.
/// A curve that only touches the other figure (y = x² and the x-axis) is found where the
/// distance has a minimum of next to nothing.
/// </summary>
public static class CurveIntersections
{
    /// <summary>The first step out from the parameter, times its size on a graph</summary>
    const double FirstStep = 1e-3;

    /// <summary>Each step out a little longer than the last, so that far is reached at all</summary>
    const double Growth = 1.1;

    /// <summary>On a graph, 2e8 times the first step away</summary>
    const int MaxSteps = 250;

    /// <summary>What counts as on the other figure, relative to the size of the coordinates: a pole or a jump of the function is no crossing</summary>
    const double RootTolerance = 1e-7;

    static readonly Point Nowhere = new Point(double.NaN, double.NaN);

    /// <summary>What a crossing with a figure goes by: a side of a Bezier path stands for the path</summary>
    public static IFigure Holder(IFigure figure)
    {
        return BezierPath.PointHolder(figure);
    }

    static bool IsBezier(IFigure figure)
    {
        return figure is Bezier || figure is BezierPath;
    }

    /// <summary>What a curve can be crossed with: what has an equation for a point to satisfy</summary>
    static bool HasEquation(IFigure figure)
    {
        return (figure is ILine || figure is IEllipse || figure is FunctionGraph) && !(figure is AngleArc);
    }

    /// <summary>
    /// Of the two figures (or the paths their sides are on), the one to walk along: a Bézier
    /// curve or path, else a graph; null when the two are not crossed this way (two lines,
    /// a line and a circle: <see cref="IntersectionPoint.GetAlgorithms"/>)
    /// </summary>
    public static IFigure FindCurve(IFigure first, IFigure second)
    {
        first = Holder(first);
        second = Holder(second);
        if (first == null || second == null || first == second)
        {
            return null;
        }

        if (IsBezier(first) && HasEquation(second))
        {
            return first;
        }

        if (IsBezier(second) && HasEquation(first))
        {
            return second;
        }

        if (first is FunctionGraph && HasEquation(second))
        {
            return first;
        }

        if (second is FunctionGraph && HasEquation(first))
        {
            return second;
        }

        return null;
    }

    public static bool CanIntersect(IFigure first, IFigure second)
    {
        return FindCurve(first, second) != null;
    }

    /// <summary>The figure that is not the curve: its equation is what the curve must satisfy</summary>
    static IFigure Other(IFigure curve, IFigure first, IFigure second)
    {
        return Holder(first) == curve ? Holder(second) : Holder(first);
    }

    /// <summary>Where on the curve the coordinates are, to look for the crossing nearest them</summary>
    public static double ParameterNear(IFigure first, IFigure second, Point coordinates)
    {
        switch (FindCurve(first, second))
        {
            case FunctionGraph:
                return coordinates.X;
            case Bezier bezier:
                return bezier.GetNearestParameterFromPoint(coordinates);
            case BezierPath path:
                return path.GetNearestParameterFromPoint(coordinates);
            default:
                return double.NaN;
        }
    }

    /// <summary>The point of the curve at the parameter; nowhere off the curve</summary>
    static Point PointAt(IFigure curve, double parameter)
    {
        switch (curve)
        {
            case FunctionGraph graph:
                return graph.GetPointFromParameter(parameter);
            case Bezier bezier:
                return parameter >= 0 && parameter <= 1 ? bezier.GetCurvePoint(parameter) : Nowhere;
            case BezierPath path:
                return path.GetPointFromParameter(parameter);
            default:
                return Nowhere;
        }
    }

    /// <summary>
    /// Zero on the figure, of one sign on one side and of the other on the other: the
    /// distance from a line or a circle, near the distance from an ellipse, the height above
    /// a graph
    /// </summary>
    static double Equation(IFigure figure, Point point)
    {
        switch (figure)
        {
            case ILine line:
            {
                var ends = line.Coordinates;
                double dx = ends.P2.X - ends.P1.X;
                double dy = ends.P2.Y - ends.P1.Y;
                double length = System.Math.Sqrt(dx * dx + dy * dy);
                if (!(length > 0))
                {
                    return double.NaN;
                }

                return ((point.X - ends.P1.X) * dy - (point.Y - ends.P1.Y) * dx) / length;
            }

            case ICircle circle:
                return circle.Radius > 0 ? circle.Center.Distance(point) - circle.Radius : double.NaN;
            case IEllipse ellipse:
            {
                double a = ellipse.SemiMajor;
                double b = ellipse.SemiMinor;
                if (!(a > 0 && b > 0))
                {
                    return double.NaN;
                }

                var center = ellipse.Center;
                double cos = System.Math.Cos(ellipse.Inclination);
                double sin = System.Math.Sin(ellipse.Inclination);
                double dx = point.X - center.X;
                double dy = point.Y - center.Y;
                double u = (dx * cos + dy * sin) / a;
                double v = (-dx * sin + dy * cos) / b;
                return (System.Math.Sqrt(u * u + v * v) - 1) * System.Math.Min(a, b);
            }

            case FunctionGraph graph:
                return point.Y - graph.GetPointFromParameter(point.X).Y;
            default:
                return double.NaN;
        }
    }

    /// <summary>
    /// The crossing of the two figures nearest the parameter on the curve, and its parameter;
    /// nowhere (the parameter as it was) when there is none. Whether it is on a ray, a segment
    /// or an arc and not only on its line or circle is the caller's to ask.
    /// </summary>
    public static Point Find(IFigure first, IFigure second, ref double parameter)
    {
        var curve = FindCurve(first, second);
        if (curve == null)
        {
            return Nowhere;
        }

        var other = Other(curve, first, second);
        double low = double.NegativeInfinity;
        double high = double.PositiveInfinity;
        if (curve is ILinearFigure linear && !(curve is FunctionGraph))
        {
            var domain = linear.GetParameterDomain();
            low = domain.Item1;
            high = domain.Item2;
        }

        double start = parameter.IsValidValue() ? parameter : 0;
        start = System.Math.Max(low, System.Math.Min(high, start));
        double step = curve is FunctionGraph ? FirstStep * System.Math.Max(1, System.Math.Abs(start)) : FirstStep;
        var search = new RootSearch(t => Equation(other, PointAt(curve, t)), t => PointAt(curve, t));
        double root = search.FindNearest(start, step, low, high);
        if (!root.IsValidValue())
        {
            return Nowhere;
        }

        var point = PointAt(curve, root);
        if (!point.Exists())
        {
            return Nowhere;
        }

        parameter = root;
        return point;
    }

    /// <summary>Whether the crossing is on the other figure itself, not only on its line or circle (a graph has no ends)</summary>
    public static bool IsOnOther(IFigure first, IFigure second, Point point)
    {
        var curve = FindCurve(first, second);
        if (curve == null)
        {
            return false;
        }

        var other = Other(curve, first, second);
        return other is FunctionGraph || other.HitTest(point) != null;
    }

    /// <summary>A root of a function of one variable, nearest a place, by a walk outwards on both sides</summary>
    class RootSearch
    {
        readonly Func<double, double> function;
        readonly Func<double, Point> curvePoint;

        public RootSearch(Func<double, double> function, Func<double, Point> curvePoint)
        {
            this.function = function;
            this.curvePoint = curvePoint;
        }

        double Value(double t)
        {
            double value = function(t);
            return value.IsValidValue() ? value : double.NaN;
        }

        /// <summary>Near enough to zero to be on both figures, for the size of the numbers there</summary>
        bool IsRoot(double t)
        {
            double value = Value(t);
            if (double.IsNaN(value))
            {
                return false;
            }

            var point = curvePoint(t);
            double scale = 1 + System.Math.Abs(point.X) + System.Math.Abs(point.Y);
            return System.Math.Abs(value) <= RootTolerance * scale;
        }

        /// <summary>A side of the walk: the last two samples</summary>
        class Side
        {
            public double Direction;
            public double PreviousT = double.NaN;
            public double PreviousValue = double.NaN;
            public double LastT;
            public double LastValue;
            public bool Done;
        }

        public double FindNearest(double start, double step, double low, double high)
        {
            double startValue = Value(start);
            if (startValue == 0)
            {
                return start;
            }

            var right = new Side { Direction = 1, LastT = start, LastValue = startValue };
            var left = new Side { Direction = -1, LastT = start, LastValue = startValue };
            double distance = 0;
            for (int i = 0; i < MaxSteps && !(right.Done && left.Done); i++)
            {
                distance += step;
                step *= Growth;
                double best = double.NaN;
                Consider(ref best, start, Advance(right, start + distance, low, high));
                Consider(ref best, start, Advance(left, start - distance, low, high));
                if (i == 0)
                {
                    // the start itself the lowest of the three: a touch there, or the
                    // crossing found last time
                    Consider(ref best, start, Touch(left.LastT, left.LastValue, start, startValue, right.LastT, right.LastValue));
                }

                if (!double.IsNaN(best))
                {
                    return best;
                }
            }

            return double.NaN;
        }

        static void Consider(ref double best, double start, double candidate)
        {
            if (!double.IsNaN(candidate) && (double.IsNaN(best) || System.Math.Abs(candidate - start) < System.Math.Abs(best - start)))
            {
                best = candidate;
            }
        }

        /// <summary>The next sample of the side, and a root between it and the samples before, nearest the start</summary>
        double Advance(Side side, double t, double low, double high)
        {
            if (side.Done)
            {
                return double.NaN;
            }

            if (t <= low || t >= high)
            {
                t = t <= low ? low : high;
                side.Done = true;
                if (t == side.LastT)
                {
                    return double.NaN;
                }
            }

            double value = Value(t);
            double result = double.NaN;
            bool lastValid = !double.IsNaN(side.LastValue);
            bool valid = !double.IsNaN(value);
            if (valid && value == 0)
            {
                result = t;
            }
            else if (lastValid && valid)
            {
                if ((side.LastValue < 0) != (value < 0))
                {
                    result = Halve(side.LastT, side.LastValue, t, value);
                }
                else
                {
                    result = Touch(side.PreviousT, side.PreviousValue, side.LastT, side.LastValue, t, value);
                }
            }
            else if (lastValid != valid)
            {
                // the function stops having a value in between: sqrt(x) meets the x-axis
                // at its very end
                double edge = valid ? FindEdge(t, side.LastT) : FindEdge(side.LastT, t);
                if (IsRoot(edge))
                {
                    result = edge;
                }
            }

            side.PreviousT = side.LastT;
            side.PreviousValue = side.LastValue;
            side.LastT = t;
            side.LastValue = value;
            return result;
        }

        /// <summary>Between two values of opposite signs: where the function is zero, or nothing for a pole or a jump</summary>
        double Halve(double a, double valueA, double b, double valueB)
        {
            for (int i = 0; i < 200; i++)
            {
                double middle = (a + b) / 2;
                if (middle == a || middle == b)
                {
                    break;
                }

                double value = Value(middle);
                if (double.IsNaN(value))
                {
                    return double.NaN;
                }

                if (value == 0)
                {
                    return middle;
                }

                if ((value < 0) == (valueA < 0))
                {
                    a = middle;
                    valueA = value;
                }
                else
                {
                    b = middle;
                    valueB = value;
                }
            }

            double root = System.Math.Abs(valueA) <= System.Math.Abs(valueB) ? a : b;
            return IsRoot(root) ? root : double.NaN;
        }

        /// <summary>
        /// Three samples of one sign, the middle one nearest zero: the lowest place of the
        /// distance between the outer two, a root if it is next to nothing there (a curve
        /// that touches the figure without crossing it)
        /// </summary>
        double Touch(double a, double valueA, double middle, double valueMiddle, double b, double valueB)
        {
            if (double.IsNaN(valueA) || double.IsNaN(valueMiddle) || double.IsNaN(valueB)
                || (valueA < 0) != (valueMiddle < 0) || (valueB < 0) != (valueMiddle < 0)
                || !(System.Math.Abs(valueMiddle) < System.Math.Abs(valueA))
                || !(System.Math.Abs(valueMiddle) <= System.Math.Abs(valueB)))
            {
                return double.NaN;
            }

            if (a > b)
            {
                (a, b) = (b, a);
            }

            // golden section on the size of the value
            const double ratio = 0.6180339887498949;
            double x1 = b - ratio * (b - a);
            double x2 = a + ratio * (b - a);
            double f1 = Size(x1);
            double f2 = Size(x2);
            for (int i = 0; i < 100 && x1 < x2; i++)
            {
                if (f1 <= f2)
                {
                    b = x2;
                    x2 = x1;
                    f2 = f1;
                    x1 = b - ratio * (b - a);
                    f1 = Size(x1);
                }
                else
                {
                    a = x1;
                    x1 = x2;
                    f1 = f2;
                    x2 = a + ratio * (b - a);
                    f2 = Size(x2);
                }
            }

            double lowest = f1 <= f2 ? x1 : x2;
            return IsRoot(lowest) ? lowest : double.NaN;
        }

        double Size(double t)
        {
            double value = Value(t);
            return double.IsNaN(value) ? double.PositiveInfinity : System.Math.Abs(value);
        }

        /// <summary>Between a place where the function has a value and one where it has none: the last place where it has</summary>
        double FindEdge(double inside, double outside)
        {
            for (int i = 0; i < 100; i++)
            {
                double middle = (inside + outside) / 2;
                if (middle == inside || middle == outside)
                {
                    break;
                }

                if (double.IsNaN(Value(middle)))
                {
                    outside = middle;
                }
                else
                {
                    inside = middle;
                }
            }

            return inside;
        }
    }
}
