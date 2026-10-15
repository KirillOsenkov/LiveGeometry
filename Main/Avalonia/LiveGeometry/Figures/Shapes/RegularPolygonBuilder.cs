using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DynamicGeometry;

/// <summary>
/// What the Square and the regular polygon tools make, as plain figures: the vertices the
/// polygon adds, a segment along each side the drawing doesn't draw already
/// (<see cref="IFigureListExtensions.FindLine"/>, as the Triangle and Polygon tools draw
/// them), and a polygon on all the vertices. A vertex is a point by coordinates written
/// over the two points the polygon is built on, named like any point, which expressions can
/// say and figures can be built on. Built on two adjacent vertices, the polygon is to the
/// left of the first to the second (a square drawn left to right stands on its side); built
/// on its center and a vertex, the vertices follow that one counterclockwise.
/// </summary>
public static class RegularPolygonBuilder
{
    public const int MinimumSides = 3;
    public const int MaximumSides = 500;

    /// <summary>
    /// Everything the polygon adds, none of it in the drawing yet, in the order to add it:
    /// the new vertices, the segments of the sides not drawn already, the polygon last. What
    /// the drawing has already is taken: a vertex that is the same point as one it shows
    /// (<see cref="FindSamePoint"/>: in a lattice of squares the fourth corner of a square is
    /// the third of its neighbor), the sides drawn between such vertices, and a polygon on
    /// the same vertices - then <paramref name="polygon"/> is that one, and none is added.
    /// </summary>
    /// <param name="onSide">Whether the two points are adjacent vertices; else the center, then a vertex</param>
    public static List<IFigure> Create(Drawing drawing, IPoint first, IPoint second, int sides, bool onSide, out Polygon polygon)
    {
        var forms = new LinearForms();
        var vertices = new List<IPoint>(onSide ? new[] { first, second } : new[] { second });
        var result = new List<IFigure>();
        foreach (var vertex in CreateVertices(drawing, first, second, sides, onSide))
        {
            var existing = FindSamePoint(drawing, (PointByCoordinates)vertex, forms);
            vertices.Add(existing ?? vertex);
            if (existing == null)
            {
                result.Add(vertex);
            }
        }

        for (int i = 0; i < vertices.Count; i++)
        {
            var from = vertices[i];
            var to = vertices[(i + 1) % vertices.Count];
            if (drawing.Figures.FindLine(from, to) == null)
            {
                result.Add(Factory.CreateSegment(drawing, from, to));
            }
        }

        polygon = drawing.Figures.OfType<Polygon>().FirstOrDefault(existing => GoesRound(existing.Dependencies, vertices));
        if (polygon == null)
        {
            polygon = Factory.CreatePolygon(drawing, vertices.Cast<IFigure>().ToList());
            result.Add(polygon);
        }

        return result;
    }

    /// <summary>
    /// A point the drawing shows where the new vertex is, that is the same point however the
    /// figures move: its X and Y, worked out exactly through the points by coordinates they
    /// name, are the new vertex's (<see cref="LinearForms"/>) - a corner of a neighbor, or the
    /// free point a fan of squares around a point closes on. A point there by chance is not
    /// (a free point the vertex isn't built to be, one typed there, a corner of a square on
    /// other points).
    /// </summary>
    static IPoint FindSamePoint(Drawing drawing, PointByCoordinates vertex, LinearForms forms)
    {
        vertex.Recalculate();
        var place = vertex.Coordinates;
        if (!vertex.Exists)
        {
            return null;
        }

        double tolerance = Math.TangencyTolerance(System.Math.Max(System.Math.Abs(place.X), System.Math.Abs(place.Y)));
        return drawing.Figures
            .OfType<IPoint>()
            .FirstOrDefault(candidate => candidate.Visible
                && candidate.Exists
                && System.Math.Abs(candidate.Coordinates.X - place.X) <= tolerance
                && System.Math.Abs(candidate.Coordinates.Y - place.Y) <= tolerance
                && forms.AreSamePoint(vertex, candidate));
    }

    /// <summary>Whether the polygon's vertices are these, going round from any of them, either way</summary>
    static bool GoesRound(IList<IFigure> polygon, List<IPoint> vertices)
    {
        int count = vertices.Count;
        int start = polygon.Count == count ? polygon.IndexOf(vertices[0]) : -1;
        if (start < 0)
        {
            return false;
        }

        bool forward = true;
        bool backward = true;
        for (int i = 0; i < count; i++)
        {
            forward &= polygon[(start + i) % count] == vertices[i];
            backward &= polygon[(start - i + count) % count] == vertices[i];
        }

        return forward || backward;
    }

    /// <summary>
    /// The vertices the polygon adds - after the two given, or after the vertex given with
    /// the center - as points by coordinates not in the drawing yet. A vertex is P + R(Q - P)
    /// for the two points P and Q, R the turn by an angle whose cosine and sine are (c, s):
    /// for a center P, the central angle times the vertex's place after Q; for a side PQ,
    /// the sums of those of the sides before the vertex, each turned by the central angle
    /// from the last.
    /// </summary>
    public static List<IPoint> CreateVertices(Drawing drawing, IPoint first, IPoint second, int sides, bool onSide)
    {
        var result = new List<IPoint>();
        double step = 2 * System.Math.PI / sides;
        int root = ExactRoot(sides);
        double cosines = 1;
        double sines = 0;
        for (int k = onSide ? 2 : 1; k < sides; k++)
        {
            double c, s;
            if (onSide)
            {
                cosines += System.Math.Cos((k - 1) * step);
                sines += System.Math.Sin((k - 1) * step);
                c = cosines;
                s = sines;
            }
            else
            {
                c = System.Math.Cos(k * step);
                s = System.Math.Sin(k * step);
            }

            var x = Formula(root, (first, "X", 1 - c), (first, "Y", s), (second, "X", c), (second, "Y", -s));
            var y = Formula(root, (first, "X", -s), (first, "Y", 1 - c), (second, "X", s), (second, "Y", c));
            result.Add(Factory.CreatePointByCoordinates(drawing, x, y));
        }

        return result;
    }

    /// <summary>
    /// The square root a polygon's coefficients are exact with: those that tile the plane -
    /// 3, 6 and 12 sides with sqrt(3), 8 with sqrt(2) (the octagons of the 4.8.8 tiling) -
    /// have nothing but halves and halves of it, a square nothing but whole numbers (0: no
    /// root). Every other polygon's are written in decimals.
    /// </summary>
    static int ExactRoot(int sides)
    {
        switch (sides)
        {
            case 3:
            case 6:
            case 12:
                return 3;
            case 8:
                return 2;
            default:
                return 0;
        }
    }

    /// <summary>
    /// "A.Y + B.X - B.Y": the terms sorted by point name, X before Y, without parentheses, a
    /// coefficient of 1 left out and one of 0 the term with it; "0" when nothing is left
    /// </summary>
    static string Formula(int root, params (IPoint Point, string Axis, double Coefficient)[] terms)
    {
        var sb = new StringBuilder();
        foreach (var term in terms
            .OrderBy(t => t.Point.Name, StringComparer.Ordinal)
            .ThenBy(t => t.Axis, StringComparer.Ordinal))
        {
            var (sign, size) = Describe(term.Coefficient, root);
            if (sign == 0)
            {
                continue;
            }

            if (sb.Length == 0)
            {
                sb.Append(sign < 0 ? "-" : "");
            }
            else
            {
                sb.Append(sign < 0 ? " - " : " + ");
            }

            if (size.Length > 0)
            {
                sb.Append(size).Append(" * ");
            }

            sb.Append(Scanner.NameInExpression(term.Point.Name)).Append('.').Append(term.Axis);
        }

        return sb.Length == 0 ? "0" : sb.ToString();
    }

    /// <summary>
    /// The coefficient's sign (0 for none) and its size as a formula writes it before a
    /// coordinate: "" for 1, "2", "0.5", "sqrt(3) / 2", "(1 + sqrt(3) / 2)" exactly where it
    /// is that (<see cref="TryExact"/>), else its value to 15 digits, never with an exponent
    /// (the expression language reads none)
    /// </summary>
    static (int Sign, string Size) Describe(double value, int root)
    {
        if (!TryExact(value, root, out int halves, out int rootHalves))
        {
            var size = ((decimal)System.Math.Abs(value)).ToString(CultureInfo.InvariantCulture);
            return (System.Math.Sign(value), size == "1" ? "" : size);
        }

        if (halves == 0 && rootHalves == 0)
        {
            return (0, null);
        }

        int sign = System.Math.Sign(halves + rootHalves * System.Math.Sqrt(root));
        halves *= sign;
        rootHalves *= sign;
        if (rootHalves == 0)
        {
            var rational = Halves(halves);
            return (sign, rational == "1" ? "" : rational);
        }

        var irrational = RootHalves(System.Math.Abs(rootHalves), root);
        if (halves == 0)
        {
            return (sign, irrational);
        }

        return (sign, "(" + Halves(halves) + (rootHalves < 0 ? " - " : " + ") + irrational + ")");
    }

    /// <summary>
    /// Whether the value is (halves + rootHalves sqrt(root)) / 2 for whole numbers no bigger
    /// than the 12-gon's coefficients need: the cosines and sines of the polygons with exact
    /// coefficients (<see cref="ExactRoot"/>) and their sums are, and a double computed from
    /// them is within rounding of it. Any number is checked for halves alone (a sum of
    /// cosines of a decagon that is 1).
    /// </summary>
    static bool TryExact(double value, int root, out int halves, out int rootHalves)
    {
        double rootValue = System.Math.Sqrt(root);
        int tries = root == 0 ? 1 : 49;
        for (int i = 0; i < tries; i++)
        {
            // 0, 1, -1, 2, -2...
            int r = (i + 1) / 2 * (i % 2 == 1 ? 1 : -1);
            double rest = 2 * value - r * rootValue;
            double whole = System.Math.Round(rest);
            if (System.Math.Abs(rest - whole) < 1e-9)
            {
                halves = (int)whole;
                rootHalves = r;
                return true;
            }
        }

        halves = 0;
        rootHalves = 0;
        return false;
    }

    /// <summary>A number of halves as it reads: "2", "0.5", "-1.5"</summary>
    static string Halves(int halves)
    {
        return halves % 2 == 0
            ? (halves / 2).ToString(CultureInfo.InvariantCulture)
            : (halves / 2.0).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>A positive number of halves of the square root: "sqrt(3) / 2", "sqrt(3)", "3 * sqrt(3) / 2"</summary>
    static string RootHalves(int rootHalves, int root)
    {
        var sqrt = "sqrt(" + root.ToString(CultureInfo.InvariantCulture) + ")";
        if (rootHalves % 2 == 0)
        {
            int wholes = rootHalves / 2;
            return wholes == 1 ? sqrt : wholes.ToString(CultureInfo.InvariantCulture) + " * " + sqrt;
        }

        return (rootHalves == 1 ? sqrt : rootHalves.ToString(CultureInfo.InvariantCulture) + " * " + sqrt) + " / 2";
    }
}
