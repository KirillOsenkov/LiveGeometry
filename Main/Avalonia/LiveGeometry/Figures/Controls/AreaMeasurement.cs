using Avalonia;
using System.Linq;

namespace LiveGeometry;

public class AreaMeasurement : Measurement
{
    protected override string Kind
    {
        get
        {
            return "Area";
        }
    }

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

    public override Point Anchor
    {
        get
        {
            return Origin;
        }
    }

    public override void UpdateVisual()
    {
        base.UpdateVisual();
        Text = Math.Round(Measure, DecimalsToShow).ToString();
    }

    private Point Origin
    {
        get
        {
            if (Dependencies[0] is PointBase)
            {
                return Dependencies.ToPoints().Midpoint();
            }
            else
            {
                return Dependencies[0].Center;
            }               
        }
    }
}
