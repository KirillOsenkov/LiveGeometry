using System.Xml;
using System.Xml.Linq;
using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// A closed shape with a length around it: a polygon (its sides added up), a circle or an
/// ellipse (the circumference), a sector or a circular segment (the arc with the two radii
/// or the chord), a closed Bezier path. NaN where there is none right now (an open path).
/// Not <see cref="ILengthProvider"/>, on purpose: a circle's Length is its radius
/// (<see cref="IFixableLength"/>), a regular polygon's its side, and an expression reads
/// c.Length as that. The perimeter goes by a name of its own, c.Perimeter.
/// </summary>
public interface IPerimeter : IFigure
{
    double Perimeter { get; }
}

/// <summary>
/// The perimeter of a shape, written beside it (the Perimeter tool). A length
/// (<see cref="ILengthProvider"/>), so that a circle's circumference can be a radius or a
/// distance to translate by. Shows the bare number, as Area and Distance do, unless a
/// <see cref="Prefix"/> is typed ("P = "): next to an Area label the two can be the same
/// number (a circle of radius 2).
/// </summary>
public class PerimeterMeasurement : Measurement, ILengthProvider
{
    /// <summary>Between the outline and the number, in pixels</summary>
    public static double Gap = 6;

    /// <summary>How far past the edge of the window a point still counts as on screen, in pixels</summary>
    public static double EdgeInset = 24;

    bool placed;

    string prefix = "";

    IPerimeter Measured
    {
        get { return Dependencies.Count > 0 ? Dependencies[0] as IPerimeter : null; }
    }

    protected override string Kind
    {
        get { return "Perimeter"; }
    }

    /// <summary>"of circle c", "of triangle ABC": "Perimeter of circle c"</summary>
    public override string Construction
    {
        get { return Measured != null ? "of " + ConstructionText.Of(Measured) : null; }
    }

    /// <summary>Text in front of the number ("P = "); nothing by default</summary>
    [PropertyGridVisible]
    [PropertyGridName("Prefix")]
    public string Prefix
    {
        get
        {
            return prefix;
        }
        set
        {
            prefix = value ?? "";
            UpdateVisual();
        }
    }

    /// <summary>
    /// A shape with a perimeter right now. Anything else has nothing to measure (an open
    /// path, a sector a file turned into an arc): the measurement then doesn't exist.
    /// </summary>
    bool HasSomethingToMeasure
    {
        get { return Measured != null && Measured.Perimeter.IsValidValue(); }
    }

    public override void UpdateExistence()
    {
        base.UpdateExistence();
        if (Exists && !HasSomethingToMeasure)
        {
            Exists = false;
        }
    }

    public double Length
    {
        get { return HasSomethingToMeasure ? Measured.Perimeter : double.NaN; }
    }

    /// <summary>
    /// A point of the outline, which is what is measured: the middle of a polygon's first
    /// side, the lower right of a circle or an ellipse (its name sits upper left), the middle
    /// of a sector's or segment's arc, the middle of a path's first piece. A circle or an
    /// ellipse too big for the window has that point off screen: then the point of it nearest
    /// the middle of the window, which moves as the view pans, as a line's name does
    /// (<see cref="FigureLabel"/>). On screen the fixed point holds, or the number would slide
    /// around a small circle at every pan.
    /// </summary>
    public override Point Anchor
    {
        get
        {
            switch (Measured)
            {
                case IPolygonalChain polygon:
                    {
                        var vertices = polygon.VertexCoordinates;
                        return vertices != null && vertices.Length >= 2
                            ? Math.Midpoint(vertices[0], vertices[1])
                            : Measured.Center;
                    }

                case EllipseArcBase arc:
                    return arc.ArcMiddle;
                case EllipseBase ellipse:
                    return OnScreenPoint(ellipse);
                case BezierPath path:
                    return path.GetPointFromParameter(0.5);
                default:
                    return Measured != null ? Measured.Center : Math.InfinitePoint;
            }
        }
    }

    Point OnScreenPoint(EllipseBase ellipse)
    {
        var center = ellipse.Center;
        var fixedPoint = Math.PointOnEllipse(center, ellipse.SemiMajor, ellipse.SemiMinor, ellipse.Inclination, -Math.PI / 4);
        if (Drawing == null || IsOnScreen(ToPhysical(fixedPoint)))
        {
            return fixedPoint;
        }

        var canvas = Drawing.CoordinateSystem.PhysicalSize;
        var middle = ToLogical(new Point(canvas.X / 2, canvas.Y / 2));
        if (!middle.Exists() || middle.Distance(center) < 1e-9)
        {
            return fixedPoint;
        }

        var crossings = Math.GetIntersectionOfEllipseAndLine(
            center,
            ellipse.SemiMajor,
            ellipse.SemiMinor,
            ellipse.Inclination,
            new PointPair(center, middle));
        if (!crossings.P1.Exists() || !crossings.P2.Exists())
        {
            return fixedPoint;
        }

        return crossings.P1.Distance(middle) <= crossings.P2.Distance(middle) ? crossings.P1 : crossings.P2;
    }

    bool IsOnScreen(Point pixel)
    {
        var canvas = Drawing.CoordinateSystem.PhysicalSize;
        return pixel.Exists()
            && pixel.X >= -EdgeInset && pixel.X <= canvas.X + EdgeInset
            && pixel.Y >= -EdgeInset && pixel.Y <= canvas.Y + EdgeInset;
    }

    public override void UpdateVisual()
    {
        if (!HasSomethingToMeasure)
        {
            return;
        }

        Text = prefix + Math.Round(Length, DecimalsToShow).ToString();

        // (placed once the window has a size: before the first layout the anchor of a
        // circle could only be the fallback, and the offset worked out from it was wrong)
        if (!placed && Drawing != null && Drawing.CoordinateSystem.PhysicalSize.X > 0)
        {
            placed = true;
            Offset = DefaultOffset();
        }

        base.UpdateVisual();
    }

    /// <summary>
    /// Where a new number goes, in pixels from the anchor: just outside the outline, clear
    /// of it - across a polygon's side away from the inside, and for the rest away from the
    /// shape's center.
    /// </summary>
    Point DefaultOffset()
    {
        var size = MeasureSize();
        var direction = OutwardDirection();
        double distance = Gap + System.Math.Abs(direction.X) * size.Width / 2 + System.Math.Abs(direction.Y) * size.Height / 2;
        return direction * distance - new Point(size.Width / 2, size.Height / 2);
    }

    /// <summary>A unit vector in pixels pointing from the anchor out of the shape</summary>
    Point OutwardDirection()
    {
        var anchor = ToPhysical(Anchor);
        var center = ToPhysical(Measured.Center);
        if (Measured is IPolygonalChain polygon && polygon.VertexCoordinates != null && polygon.VertexCoordinates.Length >= 2)
        {
            var vertices = polygon.VertexCoordinates;
            var along = RightAngleMark.Direction(ToPhysical(vertices[0]), ToPhysical(vertices[1]));
            if (along != null)
            {
                var normal = new Point(along.Value.Y, -along.Value.X);
                if ((anchor.X - center.X) * normal.X + (anchor.Y - center.Y) * normal.Y < 0)
                {
                    normal = new Point(-normal.X, -normal.Y);
                }

                return normal;
            }
        }

        return RightAngleMark.Direction(center, anchor) ?? new Point(0, 1);
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);

        // (a file written by hand without an offset gets the default place)
        placed = element.Attribute("OffsetX") != null;
        prefix = element.ReadString("Prefix") ?? "";
    }

    public override void WriteXml(XmlWriter writer)
    {
        base.WriteXml(writer);
        if (prefix.Length > 0)
        {
            writer.WriteAttributeString("Prefix", prefix);
        }
    }
}
