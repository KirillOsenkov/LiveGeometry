using System.Linq;
using System.Xml.Linq;
using Avalonia;

namespace LiveGeometry;

public class AreaMeasurement : Measurement
{
    protected override string Kind => "Area";

    /// <summary>
    /// What the Area tool takes a click on: a polygon, a circle, an ellipse, a sector or a
    /// circular segment, and a Bezier path while it has an inside (closed or filled)
    /// </summary>
    public static bool IsMeasurable(IFigure figure)
    {
        return figure is IShapeWithInterior || figure is BezierPath path && path.HasInside;
    }

    /// <summary>"of triangle ABC", "of circle c", "of ABCD" (points)</summary>
    public override string Construction
    {
        get
        {
            if (Dependencies.Count == 0)
            {
                return null;
            }

            return "of " + (Dependencies[0] is IPoint
                ? ConstructionText.Points(Dependencies.ToArray())
                : ConstructionText.Of(Dependencies[0]));
        }
    }

    public double Measure
    {
        get
        {
            switch (Dependencies[0])
            {
                case IShapeWithInterior shape:
                    return shape.Area;

                // the inside as drawn, holes left out; NaN for an open, unfilled path
                case BezierPath path:
                    return path.Area;
                default:
                    return Dependencies.ToPoints().Area();
            }
        }
    }

    /// <summary>Not while there is no area: a path opened and unfilled after it was measured (it said NaN)</summary>
    public override void UpdateExistence()
    {
        base.UpdateExistence();
        if (Exists && Dependencies.Count > 0 && !Measure.IsValidValue())
        {
            Exists = false;
        }
    }

    public override Point Anchor => Origin;

    // has its place: centered on the anchor when its number first showed, or the file's
    bool placed;

    public override void UpdateVisual()
    {
        base.UpdateVisual();
        Text = Math.Round(Measure, DecimalsToShow).ToString();
        if (!placed && HasCanvas && !Drawing.IsReading && !Text.IsEmpty())
        {
            // a new one: the number in the middle of the region (it hung its corner there)
            placed = true;
            var size = MeasureSize();
            Offset = new Point(-size.Width / 2, -size.Height / 2);
            base.UpdateVisual();
        }
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);
        placed = element.Attribute("OffsetX") != null || element.Attribute("OffsetY") != null;
    }

    /// <summary>
    /// The middle of the region: a polygon's centroid by area (<see cref="FigureAnchor.Centroid"/>:
    /// the average of the vertices lies outside an L), the center of a circle, an ellipse, a
    /// sector or a path
    /// </summary>
    private Point Origin
    {
        get
        {
            if (Dependencies[0] is PointBase)
            {
                return FigureAnchor.Centroid(Dependencies.ToPoints());
            }

            if (Dependencies[0] is PolygonBase polygon)
            {
                return FigureAnchor.Centroid(polygon.VertexCoordinates);
            }

            return Dependencies[0].Center;
        }
    }
}
