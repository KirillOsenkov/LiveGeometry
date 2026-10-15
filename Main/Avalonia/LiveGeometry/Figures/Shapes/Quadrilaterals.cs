using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// What a quadrilateral is, from where its vertices are: the most special name that fits,
/// in the US sense (a square is also a rectangle and a rhombus; a trapezoid has a pair of
/// parallel sides and isn't a parallelogram).
/// </summary>
/// <remarks>
/// Only rounding is forgiven (<see cref="Tolerance"/>), not the eye: a shape is called a square
/// when it is one - built as one, or with its vertices on the grid - and a quadrilateral dragged
/// to look like one stays a quadrilateral, as it would be in class. A looser tolerance would also
/// make the name flicker while a vertex is dragged past the special position.
/// </remarks>
public static class Quadrilaterals
{
    /// <summary>Relative: of the size of the numbers for lengths, the sine of the angle for directions</summary>
    public const double Tolerance = 1e-9;

    public static string Classify(Point a, Point b, Point c, Point d)
    {
        if (!a.Exists() || !b.Exists() || !c.Exists() || !d.Exists())
        {
            return "Quadrilateral";
        }

        var ab = b.Minus(a);
        var bc = c.Minus(b);
        var cd = d.Minus(c);
        var da = a.Minus(d);
        var scale = System.Math.Max(
            System.Math.Max(Magnitude(a), Magnitude(b)),
            System.Math.Max(Magnitude(c), Magnitude(d)));

        // a vertex on a neighbor's side, or sides that cross (a bow tie): no special name
        if (IsZero(ab, scale) || IsZero(bc, scale) || IsZero(cd, scale) || IsZero(da, scale)
            || AreParallel(ab, bc) || AreParallel(bc, cd) || AreParallel(cd, da) || AreParallel(da, ab)
            || Cross(a, b, c, d) || Cross(b, c, d, a))
        {
            return "Quadrilateral";
        }

        double sideAB = Length(ab);
        double sideBC = Length(bc);
        double sideCD = Length(cd);
        double sideDA = Length(da);

        // opposite sides equal and parallel: AB and DC the same vector
        if (IsZero(ab.Plus(cd), scale))
        {
            bool right = System.Math.Abs(Dot(ab, bc)) <= Tolerance * sideAB * sideBC;
            bool equalSides = AreEqual(sideAB, sideBC, scale);
            if (right && equalSides)
            {
                return "Square";
            }

            if (right)
            {
                return "Rectangle";
            }

            return equalSides ? "Rhombus" : "Parallelogram";
        }

        if (AreEqual(sideAB, sideBC, scale) && AreEqual(sideCD, sideDA, scale)
            || AreEqual(sideBC, sideCD, scale) && AreEqual(sideDA, sideAB, scale))
        {
            return "Kite";
        }

        if (AreParallel(ab, cd) || AreParallel(bc, da))
        {
            return "Trapezoid";
        }

        return "Quadrilateral";
    }

    static bool AreParallel(Point u, Point v)
    {
        return System.Math.Abs(u.X * v.Y - u.Y * v.X) <= Tolerance * Length(u) * Length(v);
    }

    static bool AreEqual(double x, double y, double scale)
    {
        return System.Math.Abs(x - y) <= Math.TangencyTolerance(System.Math.Max(scale, System.Math.Max(x, y)));
    }

    static bool IsZero(Point vector, double scale)
    {
        return Length(vector) <= Math.TangencyTolerance(scale);
    }

    /// <summary>Do segments PQ and RS cross each other (not just touch)?</summary>
    static bool Cross(Point p, Point q, Point r, Point s)
    {
        return Side(p, q, r) * Side(p, q, s) < 0 && Side(r, s, p) * Side(r, s, q) < 0;
    }

    static int Side(Point from, Point to, Point point)
    {
        var cross = (to.X - from.X) * (point.Y - from.Y) - (to.Y - from.Y) * (point.X - from.X);
        return System.Math.Sign(cross);
    }

    static double Dot(Point u, Point v)
    {
        return u.X * v.X + u.Y * v.Y;
    }

    static double Length(Point vector)
    {
        return System.Math.Sqrt(vector.X * vector.X + vector.Y * vector.Y);
    }

    static double Magnitude(Point point)
    {
        return System.Math.Max(System.Math.Abs(point.X), System.Math.Abs(point.Y));
    }
}
