using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Avalonia;

namespace LiveGeometry;

public interface ILengthProvider : IFigure
{
    double Length { get; }
}

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

/// <summary>What a <see cref="LengthMeasurement"/> says of the figure it measures</summary>
public enum LengthQuantity
{
    /// <summary>Between two points</summary>
    Distance,

    /// <summary>Of a segment, a vector, a polyline (or anything else with a length a file names)</summary>
    Length,

    /// <summary>Of an arc, a sector or a circular segment: the curved part alone</summary>
    [PropertyGridName("Arc length")]
    ArcLength,

    /// <summary>Of a circle or a circular arc, sector or segment</summary>
    Radius,

    /// <summary>Of a circle</summary>
    Diameter,

    /// <summary>Of a circle or an ellipse: the length around it</summary>
    Circumference,

    /// <summary>Of a polygon, a sector, a circular segment, a closed Bezier path</summary>
    Perimeter,
}

/// <summary>
/// A length, written beside what it is of (the Length tool): the distance between two
/// points, or a length of one figure - which one is <see cref="Quantity"/>, offered as the
/// figure has them (a circle's radius, diameter or circumference; an arc's length or radius;
/// a sector's arc length, radius or perimeter). This one figure stands for what were the
/// Distance and the Perimeter tools, so that there is one place to measure a length, and the
/// number can be switched to another length of the same figure after it is made.
/// A length itself (<see cref="ILengthProvider"/>): a circumference can be a radius or a
/// distance to translate by.
/// </summary>
public class LengthMeasurement : Measurement, ILengthProvider, IConditionalProperties, IOffersValues
{
    /// <summary>Between the outline and the number, in pixels</summary>
    public static double Gap = 6;

    /// <summary>How far past the edge of the window a point still counts as on screen, in pixels</summary>
    public static double EdgeInset = 24;

    /// <summary>Null: the measured figure's first quantity (<see cref="DefaultQuantity"/>)</summary>
    LengthQuantity? quantity;

    /// <summary>Null: the quantity's own (<see cref="DefaultPrefix"/>)</summary>
    string prefix;

    /// <summary>Whether the number has its place, or gets the default one at the next redraw</summary>
    bool placed;

    public LengthMeasurement()
    {
    }

    /// <param name="quantity">Which length; null for the figure's first</param>
    public LengthMeasurement(LengthQuantity? quantity)
    {
        this.quantity = quantity;
    }

    /// <summary>The one figure measured; null for the distance between two points</summary>
    IFigure Measured
    {
        get { return Dependencies.Count == 1 ? Dependencies[0] : null; }
    }

    bool IsBetweenPoints
    {
        get { return Dependencies.Count == 2 && Dependencies[0] is IPoint && Dependencies[1] is IPoint; }
    }

    /// <summary>
    /// What the Length tool takes a click on to measure: a segment, a vector, a polyline, a
    /// Bezier curve or path, an arc of any kind, a circle, an ellipse, a polygon - whatever
    /// has an end or goes round, while it has a length. Not what goes on without end (a
    /// line, a ray, a function graph: a click there puts a point on it, for a distance), nor
    /// a locus, sampled by the window; not a number or a label (a length too, but nothing to
    /// measure), nor an angle's mark, an arc in code only, sized in pixels
    /// (<see cref="IFigureExtensions.GivesLength"/>).
    /// </summary>
    public static bool IsMeasurable(IFigure figure)
    {
        switch (figure)
        {
            case AngleArc:
                return false;
            case Segment:
            case Vector:
            case Polyline:
            case Bezier:
            case BezierPath:
            case EllipseArcBase:
            case EllipseBase:
            case PolygonBase:
                var quantities = QuantitiesOf(figure);
                return quantities.Count > 0 && ValueOf(figure, quantities[0]).IsValidValue();
            default:
                return false;
        }
    }

    /// <summary>
    /// The lengths a figure has, in the order the row offers them: the first is what a click
    /// on it measures. A sector's arc length comes first, as an arc's does: the curve is
    /// what is special about it, and the perimeter is one choice away.
    /// </summary>
    public static IReadOnlyList<LengthQuantity> QuantitiesOf(IFigure figure)
    {
        switch (figure)
        {
            case AngleArc:
                return System.Array.Empty<LengthQuantity>();
            case EllipseArcBase arc:
                {
                    var result = new List<LengthQuantity> { LengthQuantity.ArcLength };
                    if (arc is CircleArcBase)
                    {
                        result.Add(LengthQuantity.Radius);
                    }

                    if (arc is IPerimeter)
                    {
                        result.Add(LengthQuantity.Perimeter);
                    }

                    return result;
                }

            case CircleBase:
                return new[] { LengthQuantity.Radius, LengthQuantity.Diameter, LengthQuantity.Circumference };
            case EllipseBase:
                return new[] { LengthQuantity.Circumference };

            // (opened or closed after it is measured, a path's number follows: the quantity
            // is the first one unless another was picked, and a path has one either way)
            case BezierPath path:
                return new[] { path.Closed ? LengthQuantity.Perimeter : LengthQuantity.Length };
            case Bezier:
                return new[] { LengthQuantity.Length };
            case IPerimeter:
                return new[] { LengthQuantity.Perimeter };
            case ILengthProvider:
                return new[] { LengthQuantity.Length };
            default:
                return System.Array.Empty<LengthQuantity>();
        }
    }

    /// <summary>The quantities this measurement can say of what it measures</summary>
    public IReadOnlyList<LengthQuantity> Quantities
    {
        get
        {
            if (IsBetweenPoints)
            {
                return new[] { LengthQuantity.Distance };
            }

            return Measured != null ? QuantitiesOf(Measured) : System.Array.Empty<LengthQuantity>();
        }
    }

    LengthQuantity DefaultQuantity
    {
        get
        {
            var quantities = Quantities;
            return quantities.Count > 0 ? quantities[0] : LengthQuantity.Length;
        }
    }

    /// <summary>
    /// Which length of the figure the number is: its radius, its diameter, its circumference...
    /// (<see cref="Quantities"/>). Read-only where the figure has one. A new one has the
    /// figure's first; a file says it only when it is another.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridName("Measure")]
    [PropertyGridDisallowMultiEdit]
    [PropertyGridCustomValueProvider(typeof(QuantityValue))]
    public LengthQuantity Quantity
    {
        get
        {
            return quantity ?? DefaultQuantity;
        }
        set
        {
            if (Quantity == value)
            {
                return;
            }

            // a number of a new kind sits somewhere else (the middle of a radius, not of the
            // arc): placed again at the next redraw
            SetQuantityState(value, Offset, placed: false);
        }
    }

    void SetQuantityState(LengthQuantity? newQuantity, Point offset, bool placed)
    {
        quantity = newQuantity;
        Offset = offset;
        this.placed = placed;
        if (Drawing != null)
        {
            UpdateExistence();
            UpdateVisual();
            this.RecalculateAllDependents();
        }

        RaisePropertyChanged(nameof(Prefix));
    }

    /// <summary>
    /// The Measure row: undo puts back the quantity and where the number was - a new quantity
    /// places the number anew, and undo of it left the number at the new place, away from
    /// where it had been dragged to
    /// </summary>
    public class QuantityValue : ConditionalPropertyValue, IRestorableValue
    {
        record SavedQuantity(LengthQuantity? Quantity, Point Offset, bool Placed);

        LengthMeasurement Owner
        {
            get { return (LengthMeasurement)Parent; }
        }

        public object CaptureState()
        {
            var owner = Owner;
            return new SavedQuantity(owner.quantity, owner.Offset, owner.placed);
        }

        public void RestoreState(object state)
        {
            var saved = (SavedQuantity)state;
            Owner.SetQuantityState(saved.Quantity, saved.Offset, saved.Placed);
        }
    }

    /// <summary>What the number is called in the grid's title and the Figure List: "Radius of circle c"</summary>
    protected override string Kind
    {
        get { return KindOf(Quantity); }
    }

    public static string KindOf(LengthQuantity quantity)
    {
        return quantity == LengthQuantity.ArcLength ? "Arc length" : quantity.ToString();
    }

    /// <summary>"AB" for a distance ("Distance AB"), "of circle c" for the rest ("Radius of circle c")</summary>
    public override string Construction
    {
        get
        {
            if (IsBetweenPoints)
            {
                return ConstructionText.Points(Dependencies[0], Dependencies[1]);
            }

            return Measured != null ? "of " + ConstructionText.Of(Measured) : null;
        }
    }

    /// <summary>
    /// What another figure's construction calls the number when it is built on it (a circle
    /// with it for a radius): AB for a distance or a segment's length, as the segment itself
    /// would be named, "circumference of circle c" for the rest
    /// </summary>
    public string ValueText
    {
        get
        {
            if (IsBetweenPoints)
            {
                return Construction;
            }

            if (Quantity == LengthQuantity.Length && (Measured is Segment || Measured is Vector))
            {
                return NameDisplay.Format(Measured.Name);
            }

            return KindOf(Quantity).ToLowerInvariant() + " " + Construction;
        }
    }

    /// <summary>
    /// Text in front of the number: "r = ", "d = ", "C = " and "P = " for a radius, a
    /// diameter, a circumference and a perimeter (a circle of radius 2 has 12.57 for both its
    /// circumference and its area, and r = 2 next to d = 4 says which is which), nothing for
    /// a distance, a length and an arc length, which sit on what they measure. Typed, it stays
    /// what was typed whatever the quantity becomes; typed as the quantity's own, it follows
    /// the quantity again.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridName("Prefix")]
    [PropertyGridCustomValueProvider(typeof(PrefixValue))]
    public string Prefix
    {
        get
        {
            return prefix ?? DefaultPrefix(Quantity);
        }
        set
        {
            prefix = value == null || value == DefaultPrefix(Quantity) ? null : value;
            UpdateVisual();
        }
    }

    /// <summary>
    /// The Prefix row: undo puts back the prefix as it was kept, typed or the quantity's own.
    /// Through the setter, a typed "d = " on a radius later measured as the diameter came back
    /// as the diameter's own and no longer stayed when the quantity was undone.
    /// </summary>
    public class PrefixValue : PropertyValue, IRestorableValue
    {
        public object CaptureState()
        {
            return new[] { ((LengthMeasurement)Parent).prefix };
        }

        public void RestoreState(object state)
        {
            var owner = (LengthMeasurement)Parent;
            owner.prefix = ((string[])state)[0];
            owner.UpdateVisual();
            owner.RaisePropertyChanged(nameof(Prefix));
        }
    }

    public static string DefaultPrefix(LengthQuantity quantity)
    {
        switch (quantity)
        {
            case LengthQuantity.Radius:
                return "r = ";
            case LengthQuantity.Diameter:
                return "d = ";
            case LengthQuantity.Circumference:
                return "C = ";
            case LengthQuantity.Perimeter:
                return "P = ";
            default:
                return "";
        }
    }

    /// <summary>
    /// The quantity is one the figure has, and it has a value right now. Anything else has
    /// nothing to measure (an open path, a segment replaced by a ray, a sector a file turned
    /// into an arc): the measurement then doesn't exist, where it used to throw on every redraw.
    /// </summary>
    bool HasSomethingToMeasure
    {
        get { return Quantities.Contains(Quantity) && Length.IsValidValue(); }
    }

    public override void UpdateExistence()
    {
        base.UpdateExistence();
        if (Exists && !HasSomethingToMeasure)
        {
            Exists = false;
        }
    }

    /// <summary>The number, NaN when the figure has no such length</summary>
    public double Length
    {
        get
        {
            if (IsBetweenPoints)
            {
                return Point(0).Distance(Point(1));
            }

            return Measured != null ? ValueOf(Measured, Quantity) : double.NaN;
        }
    }

    /// <summary>The figure's length of the kind, NaN when it has none</summary>
    static double ValueOf(IFigure measured, LengthQuantity quantity)
    {
        switch (quantity)
        {
            case LengthQuantity.Length when measured is ILengthProvider withLength:
                return withLength.Length;
            case LengthQuantity.Length when measured is Bezier curve:
                return curve.Length;
            case LengthQuantity.Length when measured is BezierPath path:
                return path.Length;
            case LengthQuantity.ArcLength when measured is EllipseArcBase arc:
                return arc.Length;
            case LengthQuantity.Radius:
                return RadiusOf(measured);
            case LengthQuantity.Diameter:
                return 2 * RadiusOf(measured);
            case LengthQuantity.Circumference when measured is EllipseBase ellipse:
                return ellipse.Perimeter;
            case LengthQuantity.Perimeter when measured is IPerimeter shape:
                return shape.Perimeter;
            default:
                return double.NaN;
        }
    }

    static double RadiusOf(IFigure figure)
    {
        switch (figure)
        {
            case CircleBase circle:
                return circle.Radius;
            case CircleArcBase arc:
                return arc.Radius;
            default:
                return double.NaN;
        }
    }

    /// <summary>A number of these quantities is placed outside what it measures when it is made; a distance and a length sit at the middle, as they always did</summary>
    bool PlacesItself
    {
        get { return Quantity != LengthQuantity.Distance && Quantity != LengthQuantity.Length; }
    }

    /// <summary>
    /// What the offset is from: the middle of the two points or of the figure for a distance
    /// or a length; the middle of an arc for its length; the middle of a radius for a radius
    /// (the one to the point a circle goes through, or to an arc's start, else the one to the
    /// right); the center for a diameter; a point of the outline for a circumference or a
    /// perimeter - the middle of a polygon's first side, the lower right of a circle or an
    /// ellipse (its name sits upper left), the middle of a sector's or segment's arc, the
    /// middle of a path's first piece. A circle or an ellipse too big for the window has that
    /// point off screen: then the point of it nearest the middle of the window, which moves as
    /// the view pans, as a line's name does (<see cref="FigureLabel"/>). On screen the fixed
    /// point holds, or the number would slide around a small circle at every pan.
    /// </summary>
    public override Point Anchor
    {
        get
        {
            if (IsBetweenPoints)
            {
                return Math.Midpoint(Point(0), Point(1));
            }

            var measured = Measured;
            if (measured == null)
            {
                return Math.InfinitePoint;
            }

            switch (Quantity)
            {
                case LengthQuantity.ArcLength when measured is EllipseArcBase arc:
                    return arc.ArcMiddle;
                case LengthQuantity.Radius:
                    return RadiusMiddle(measured);
                case LengthQuantity.Diameter:
                    return measured.Center;
                case LengthQuantity.Circumference when measured is EllipseBase ellipse:
                    return OnScreenPoint(ellipse);
                case LengthQuantity.Perimeter:
                    return OutlinePoint(measured);

                // a curve's middle is on it, where a polyline's or a curve's center may be far off
                case LengthQuantity.Length when measured is Polyline || measured is Bezier || measured is BezierPath:
                    {
                        var curve = (ILinearFigure)measured;
                        var domain = curve.GetParameterDomain();
                        var middle = curve.GetPointFromParameter((domain.Item1 + domain.Item2) / 2);
                        return middle.Exists() ? middle : measured.Center;
                    }

                default:
                    return measured.Center;
            }
        }
    }

    static Point RadiusMiddle(IFigure measured)
    {
        switch (measured)
        {
            case Circle circle:
                return Math.Midpoint(circle.Center, circle.Point(1));
            case CircleArcBase arc:
                return Math.Midpoint(arc.Center, arc.BeginLocation);
            case CircleBase circle:
                return circle.Center + new Point(circle.Radius / 2, 0);
            default:
                return measured.Center;
        }
    }

    static Point OutlinePoint(IFigure measured)
    {
        switch (measured)
        {
            case IPolygonalChain polygon:
                {
                    var vertices = polygon.VertexCoordinates;
                    return vertices != null && vertices.Length >= 2
                        ? Math.Midpoint(vertices[0], vertices[1])
                        : measured.Center;
                }

            case EllipseArcBase arc:
                return arc.ArcMiddle;
            case BezierPath path:
                return path.GetPointFromParameter(0.5);
            default:
                return measured.Center;
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

        Text = Prefix + Math.Round(Length, DecimalsToShow).ToString();

        // Where the number goes is decided once, at its first redraw (or by the file, or by a
        // new quantity picked in the row), whatever its quantity: a distance or a length keeps
        // its offset, the rest get the default place. Not decided for a length then, the
        // number of an open path was placed the first time the path was closed, outside the
        // undo step and only until the file was saved (read back, it had an offset).
        // (Once the window has a size: before the first layout the anchor of a circle could
        // only be the fallback, and the offset worked out from it was wrong.)
        if (!placed && Drawing != null && Drawing.CoordinateSystem.PhysicalSize.X > 0)
        {
            placed = true;
            if (PlacesItself)
            {
                Offset = DefaultOffset();
            }
        }

        base.UpdateVisual();
    }

    /// <summary>
    /// Where a new number goes, in pixels from the anchor: clear of what it measures - across
    /// a polygon's side away from the inside, beside a radius, under a diameter's center, and
    /// for the rest away from the shape's center.
    /// </summary>
    Point DefaultOffset()
    {
        var size = MeasureSize();
        var direction = OutwardDirection();
        double distance = Gap + System.Math.Abs(direction.X) * size.Width / 2 + System.Math.Abs(direction.Y) * size.Height / 2;
        return direction * distance - new Point(size.Width / 2, size.Height / 2);
    }

    /// <summary>A unit vector in pixels pointing from the anchor to where the number goes</summary>
    Point OutwardDirection()
    {
        var measured = Measured;
        var anchor = ToPhysical(Anchor);
        var center = ToPhysical(measured.Center);
        switch (Quantity)
        {
            case LengthQuantity.Diameter:
                return new Point(0, 1);
            case LengthQuantity.Radius:
                {
                    // across the radius, on its upper side
                    var along = RightAngleMark.Direction(center, anchor);
                    if (along == null)
                    {
                        return new Point(0, -1);
                    }

                    var normal = new Point(along.Value.Y, -along.Value.X);
                    return normal.Y <= 0 ? normal : new Point(-normal.X, -normal.Y);
                }
        }

        if (measured is IPolygonalChain polygon && polygon.VertexCoordinates != null && polygon.VertexCoordinates.Length >= 2)
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

    /// <summary>The Measure row is read-only where the figure has one length</summary>
    public bool CanEdit(string propertyName)
    {
        return propertyName != nameof(Quantity) || Quantities.Count > 1;
    }

    public string Caption(string propertyName, string defaultCaption)
    {
        return defaultCaption;
    }

    /// <summary>The Measure row offers the lengths the figure has</summary>
    public bool Offers(string propertyName, object value)
    {
        return propertyName != nameof(Quantity) || value is LengthQuantity offered && Quantities.Contains(offered);
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);

        // (a file written by hand without an offset gets the default place)
        placed = element.Attribute("OffsetX") != null;
        prefix = element.ReadString("Prefix");
        quantity = System.Enum.TryParse(element.ReadString("Quantity"), out LengthQuantity read) ? read : null;
    }

    public override void WriteXml(XmlWriter writer)
    {
        base.WriteXml(writer);
        if (quantity != null && quantity != DefaultQuantity)
        {
            writer.WriteAttributeString("Quantity", quantity.ToString());
        }

        if (prefix != null)
        {
            writer.WriteAttributeString("Prefix", prefix);
        }
    }
}
