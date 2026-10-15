using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// A point kept inside a figure (<see cref="IRegion"/>: a polygon, a circle, a sector, a
/// path), GeoGebra's "Point on Object". Dragged like a free point, but never out of the
/// figure. Its place is an <see cref="Offset"/> from the figure's center, so that it is
/// worked out from the figure as it is now however the figure got there: dragged, driven
/// by a slider, the image of a transformation. Where that place falls outside (the figure
/// shrank or turned), the point sits at the nearest point inside and the offset keeps the
/// wished-for place, so undo of the change, or the figure growing back, gives the point its
/// place back. Only a drag of the point (or a file) sets the offset.
/// </summary>
public partial class PointInRegion : FreePoint, IPoint
{
    /// <summary>Where the point wants to be, from the figure's center</summary>
    public Point Offset { get; set; }

    // a hand-written file may give X and Y alone: the offset is taken from them once the
    // figure is in the drawing
    bool hasOffset;

    public IRegion Region
    {
        get
        {
            return Dependencies.Count > 0 ? Dependencies[0] as IRegion : null;
        }
    }

    // a FreePoint by inheritance, dragged within its figure only (ZOrder)
    protected override ZOrder DefaultLayer()
    {
        return ZOrder.ConstrainedPoints;
    }

    public override void ReadXml(System.Xml.Linq.XElement element)
    {
        base.ReadXml(element);
        hasOffset = element.Attribute("OffsetX") != null;
        if (hasOffset)
        {
            Offset = new Point(element.ReadDouble("OffsetX"), element.ReadDouble("OffsetY"));
        }
    }

    public override void WriteXml(System.Xml.XmlWriter writer)
    {
        base.WriteXml(writer);
        writer.WriteAttributeDouble("OffsetX", Offset.X);
        writer.WriteAttributeDouble("OffsetY", Offset.Y);
    }

    /// <summary>"inside triangle ABC"</summary>
    public override string Construction
    {
        get
        {
            return Dependencies.Count > 0 ? "inside " + ConstructionText.Of(Dependencies[0]) : null;
        }
    }

    protected override Avalonia.Controls.Shapes.Shape CreateShape()
    {
        var result = Factory.CreateDependentPointShape();
        result.Fill = new SolidColorBrush(Color.FromArgb(255, 128, 255, 128));
        return result;
    }

    public override object CapturePlace()
    {
        return Offset;
    }

    public override void RestorePlace(object place)
    {
        Offset = (Point)place;
        this.RecalculateAndUpdateVisual();
    }

    public override bool AllowMove()
    {
        return !Locked;
    }

    // Not typed, as a point on a figure's aren't: a place typed outside the region went to
    // the nearest place inside, the other coordinate changing with it. Dragged only.
    public override double X => base.X;

    public override double Y => base.Y;

    /// <summary>The point goes to the nearest place inside, and that is its place from now on</summary>
    public override void MoveToCore(Point newPosition)
    {
        var region = Region;
        if (region != null && region.Exists && region.HasInside)
        {
            var inside = region.NearestPointInside(newPosition);
            var center = region.Center;
            if (inside.Exists() && center.Exists())
            {
                newPosition = inside;
                Offset = inside.Minus(center);
                hasOffset = true;
            }
        }

        base.MoveToCore(newPosition);
    }

    public override void Recalculate()
    {
        var region = Region;
        if (region == null || !Dependencies.Exists() || !region.HasInside)
        {
            Exists = false;
            return;
        }

        var center = region.Center;
        if (!center.Exists())
        {
            Exists = false;
            return;
        }

        if (!hasOffset)
        {
            Offset = Coordinates.Minus(center);
            hasOffset = true;
        }

        var place = region.NearestPointInside(center.Plus(Offset));
        if (!place.Exists())
        {
            Exists = false;
            return;
        }

        Coordinates = place;
        Exists = true;
    }

    /// <summary>
    /// The figures a click here puts a point inside of, the smallest first (the inner square
    /// of a square in a square): every visible region containing the place.
    /// </summary>
    public static List<IRegion> FindRegions(Drawing drawing, Point coordinates)
    {
        return drawing.Figures
            .OfType<IRegion>()
            .Where(region => region.Visible
                && region.Exists
                && region.IsHitTestVisible
                && region.HasInside
                && region.ContainsPoint(coordinates))
            .OrderBy(region => System.Math.Abs(region.Area))
            .ToList();
    }

    /// <summary>Lets the point go: a free point where it is (<see cref="PointSnapping.Release"/>)</summary>
    [PropertyGridVisible]
    [PropertyGridName("Free point")]
    [PropertyGridIcon(PropertyGridIcon.Unlock)]
    public void Release()
    {
        PointSnapping.Release(this);
    }
}
