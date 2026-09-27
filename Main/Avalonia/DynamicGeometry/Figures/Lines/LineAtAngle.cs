using System.Collections.Generic;

namespace DynamicGeometry;

/// <summary>
/// A line through a point at an angle to the x axis, counterclockwise. The angle is a figure
/// it depends on: a <see cref="Number"/> holding a typed value, edited in the property grid,
/// or anything with an angle (an angle measurement or its arc, a slider in degrees), which
/// the line then follows. Dependencies: the point, then the angle.
/// </summary>
public class LineAtAngle : LineBase, ILine, IConditionalProperties
{
    public static LineAtAngle Create(Drawing drawing, IPoint point, IFigure angleSource)
    {
        return new LineAtAngle() { Drawing = drawing, Dependencies = new List<IFigure>() { point, angleSource } };
    }

    public IFigure AngleSource
    {
        get { return Dependencies.Count > 1 ? Dependencies[1] : null; }
    }

    double Radians
    {
        get { return AngleSource is IAngleProvider angle ? angle.Angle : double.NaN; }
    }

    public override PointPair Coordinates
    {
        get
        {
            var point = Point(0);
            return new PointPair(point, Math.GetTranslationPoint(point, 1, Radians));
        }
    }

    public override PointPair OnScreenCoordinates
    {
        get { return Math.GetLineFromSegment(Coordinates, CanvasLogicalBorders); }
    }

    /// <summary>An angle that isn't a number (a side of the measured angle has no length) leaves no line</summary>
    public override void UpdateExistence()
    {
        base.UpdateExistence();
        if (Exists && !Coordinates.P2.Exists())
        {
            Exists = false;
        }
    }

    /// <summary>In degrees; editable when it is a typed value, which lives in the Number</summary>
    [PropertyGridVisible]
    [PropertyGridPreferredEditor("UpDown")]
    [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
    public override double Angle
    {
        get
        {
            return AngleSource is Number number ? number.Value : Radians.ToDegrees();
        }
        set
        {
            if (AngleSource is Number number)
            {
                number.Value = value;
            }
        }
    }

    public bool CanEdit(string propertyName)
    {
        return propertyName != "Angle" || AngleSource is Number;
    }

    /// <summary>An angle taken from a figure says which</summary>
    public string Caption(string propertyName, string defaultCaption)
    {
        var source = AngleSource;
        return propertyName != "Angle" || source == null || source is Number ? defaultCaption : defaultCaption + " = " + source.Name;
    }
}
