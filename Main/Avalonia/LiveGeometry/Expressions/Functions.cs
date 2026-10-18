using Avalonia;

namespace LiveGeometry;

/// <summary>
/// The functions of the expression language, all of them (2026-10-18): a method here is
/// called by the name <see cref="FunctionCatalog"/> gives it, in any case, and by nothing
/// else - System.Math is no longer searched by name, so BitIncrement, FusedMultiplyAdd
/// and Log(a, newBase) are no longer the language. Trigonometry works in radians, as in
/// Desmos and GeoGebra (30° is a number in radians); what is read off the drawing is in
/// degrees, as the Angle tool shows it. A function of numbers takes doubles and gives a
/// double; one of points takes their coordinates (area takes the points themselves). The
/// player has the twin of each (src/expressions/functions.js).
/// </summary>
public static class Functions
{
    public static double Sin(double x)
    {
        return System.Math.Sin(x);
    }

    public static double Cos(double x)
    {
        return System.Math.Cos(x);
    }

    public static double Tan(double x)
    {
        return System.Math.Tan(x);
    }

    public static double Sec(double x)
    {
        return 1 / System.Math.Cos(x);
    }

    public static double Csc(double x)
    {
        return 1 / System.Math.Sin(x);
    }

    public static double Cot(double x)
    {
        return System.Math.Cos(x) / System.Math.Sin(x);
    }

    public static double Arcsin(double x)
    {
        return System.Math.Asin(x);
    }

    public static double Arccos(double x)
    {
        return System.Math.Acos(x);
    }

    public static double Arctan(double x)
    {
        return System.Math.Atan(x);
    }

    public static double Atan2(double y, double x)
    {
        return System.Math.Atan2(y, x);
    }

    public static double Sinh(double x)
    {
        return System.Math.Sinh(x);
    }

    public static double Cosh(double x)
    {
        return System.Math.Cosh(x);
    }

    public static double Tanh(double x)
    {
        return System.Math.Tanh(x);
    }

    public static double Arcsinh(double x)
    {
        return System.Math.Asinh(x);
    }

    public static double Arccosh(double x)
    {
        return System.Math.Acosh(x);
    }

    public static double Arctanh(double x)
    {
        return System.Math.Atanh(x);
    }

    public static double Sqrt(double x)
    {
        return System.Math.Sqrt(x);
    }

    public static double Cbrt(double x)
    {
        return System.Math.Cbrt(x);
    }

    /// <summary>The nth root; of a negative number only for an odd whole n (the cube root of -8 is -2)</summary>
    public static double Nthroot(double x, double n)
    {
        if (x < 0 && n == System.Math.Floor(n) && System.Math.Abs(n % 2) == 1)
        {
            return -System.Math.Pow(-x, 1 / n);
        }

        return System.Math.Pow(x, 1 / n);
    }

    public static double Exp(double x)
    {
        return System.Math.Exp(x);
    }

    public static double Ln(double x)
    {
        return System.Math.Log(x);
    }

    /// <summary>The common logarithm, to base 10, as a US classroom and every calculator say log</summary>
    public static double Log(double x)
    {
        return System.Math.Log10(x);
    }

    /// <summary>log_2(x), log_10(x): the base is in the function's name, as Desmos writes it</summary>
    public static double LogBase(double x, double logBase)
    {
        return System.Math.Log(x) / System.Math.Log(logBase);
    }

    public static double Abs(double x)
    {
        return System.Math.Abs(x);
    }

    /// <summary>A half goes up, as at school: System.Math.Round takes 2.5 to 2</summary>
    public static double Round(double x)
    {
        return System.Math.Round(x, System.MidpointRounding.AwayFromZero);
    }

    /// <summary>
    /// Rounded to so many digits after the point, a half up as the labels show numbers
    /// (<see cref="Math.Round(double, int)"/>); fewer than none round to tens, hundreds
    /// </summary>
    public static double Round(double x, double digits)
    {
        if (digits != System.Math.Floor(digits) || !x.IsValidValue())
        {
            return double.NaN;
        }

        if (digits < 0)
        {
            double factor = System.Math.Pow(10, -digits);
            return Round(x / factor) * factor;
        }

        return digits > 15 ? x : Math.Round(x, (int)digits);
    }

    public static double Floor(double x)
    {
        return System.Math.Floor(x);
    }

    public static double Ceil(double x)
    {
        return System.Math.Ceiling(x);
    }

    /// <summary>System.Math.Sign throws for what is not a number (the sign of log(x) left of 0)</summary>
    public static double Sign(double x)
    {
        return double.IsNaN(x) ? double.NaN : System.Math.Sign(x);
    }

    /// <summary>What is left of a over whole b's, as Desmos says it: never negative for a b above 0 (mod(-1, 3) is 2)</summary>
    public static double Mod(double a, double b)
    {
        return b == 0 ? double.NaN : a - b * System.Math.Floor(a / b);
    }

    /// <summary>The smaller of two; min(a, b, c) takes them two at a time</summary>
    public static double Min(double a, double b)
    {
        return System.Math.Min(a, b);
    }

    public static double Max(double a, double b)
    {
        return System.Math.Max(a, b);
    }

    /// <summary>System.Math.Clamp throws when the bounds are the wrong way round</summary>
    public static double Clamp(double x, double low, double high)
    {
        return low > high ? double.NaN : System.Math.Max(low, System.Math.Min(high, x));
    }

    /// <summary>The greatest common factor of two whole numbers; none of anything else</summary>
    public static double Gcd(double a, double b)
    {
        if (!IsWhole(a) || !IsWhole(b))
        {
            return double.NaN;
        }

        a = System.Math.Abs(a);
        b = System.Math.Abs(b);
        while (b != 0)
        {
            (a, b) = (b, a % b);
        }

        return a;
    }

    /// <summary>The least common multiple of two whole numbers</summary>
    public static double Lcm(double a, double b)
    {
        double gcd = Gcd(a, b);
        if (double.IsNaN(gcd))
        {
            return double.NaN;
        }

        return gcd == 0 ? 0 : System.Math.Abs(a / gcd * b);
    }

    static bool IsWhole(double x)
    {
        return x.IsValidValue() && x == System.Math.Floor(x) && System.Math.Abs(x) < 1e15;
    }

    /// <summary>An angle x in radians, in degrees: deg(pi) is 180</summary>
    public static double Deg(double x)
    {
        return x.ToDegrees();
    }

    /// <summary>An angle x in degrees, in radians: rad(180) is pi; the same as x°</summary>
    public static double Rad(double x)
    {
        return x.ToRadians();
    }

    /// <summary>a where the condition holds (is not 0), b where it doesn't, nothing where it has no value</summary>
    public static double If(double condition, double a, double b)
    {
        if (double.IsNaN(condition))
        {
            return double.NaN;
        }

        return condition != 0 ? a : b;
    }

    /// <summary>a where the condition holds, and no value where it doesn't: what is built on it isn't there</summary>
    public static double If(double condition, double a)
    {
        return If(condition, a, double.NaN);
    }

    // The comparisons: 1 where they hold and 0 where they don't. Two numbers that differ by
    // rounding alone are equal (a point dragged onto a circle, AB = CD for a rhombus), with
    // the relative tolerance by which figures touch (Math.TangencyTolerance).

    public static double Less(double a, double b)
    {
        return Compare(a, b, (x, y) => x < y && !AreEqual(x, y));
    }

    public static double LessOrEqual(double a, double b)
    {
        return Compare(a, b, (x, y) => x < y || AreEqual(x, y));
    }

    public static double Greater(double a, double b)
    {
        return Compare(a, b, (x, y) => x > y && !AreEqual(x, y));
    }

    public static double GreaterOrEqual(double a, double b)
    {
        return Compare(a, b, (x, y) => x > y || AreEqual(x, y));
    }

    public static double Equal(double a, double b)
    {
        return Compare(a, b, AreEqual);
    }

    public static double NotEqual(double a, double b)
    {
        return Compare(a, b, (x, y) => !AreEqual(x, y));
    }

    static double Compare(double a, double b, System.Func<double, double, bool> holds)
    {
        if (double.IsNaN(a) || double.IsNaN(b))
        {
            return double.NaN;
        }

        return holds(a, b) ? 1 : 0;
    }

    static bool AreEqual(double a, double b)
    {
        return a == b || System.Math.Abs(a - b) <= Math.TangencyTolerance(System.Math.Max(System.Math.Abs(a), System.Math.Abs(b)));
    }

    public static double Distance(Point a, Point b)
    {
        return a.Distance(b);
    }

    /// <summary>The angle ABC as the Angle tool measures it: the smaller angle at b, in degrees</summary>
    public static double Angle(Point a, Point b, Point c)
    {
        var a1 = Math.GetAngle(b, a);
        var a2 = Math.GetAngle(b, c);
        double result = System.Math.Abs(a2 - a1);
        if (result >= Math.PI)
        {
            result = 2 * Math.PI - result;
        }

        return result.ToDegrees();
    }

    /// <summary>The angle ABC at b counterclockwise from BA to BC, in degrees, from 0 up to 360</summary>
    public static double Directedangle(Point a, Point b, Point c)
    {
        return Math.OAngle(a, b, c).ToDegrees();
    }

    /// <summary>The direction from a to b in degrees, counterclockwise from the x axis, as a segment's Direction row says it</summary>
    public static double Direction(Point a, Point b)
    {
        return DirectionConstraint.Of(new PointPair(a, b));
    }

    public static double Area(params IPoint[] points)
    {
        return points.ToPoints().Area();
    }
}
