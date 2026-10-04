using System.Collections.Generic;

namespace DynamicGeometry;

/// <summary>
/// A line through a point at an angle to the x axis, counterclockwise. The angle is a figure
/// it depends on: a <see cref="Number"/> holding a typed value, edited in the property grid,
/// or anything with an angle (an angle measurement or its arc, a slider in degrees), which
/// the line then follows (<see cref="ITiedValues"/>). Dependencies: the point, then the angle.
/// </summary>
public class LineAtAngle : LineBase, ILine, ITiedValues, IConditionalProperties
{
    public static LineAtAngle Create(Drawing drawing, IPoint point, IFigure angleSource)
    {
        return new LineAtAngle() { Drawing = drawing, Dependencies = new List<IFigure>() { point, angleSource } };
    }

    /// <summary>
    /// An angle measurement, its arc or a slider. Not any other arc, although it has an angle
    /// too: a click on an arc puts a point on it.
    /// </summary>
    public static bool CanTakeAngleFrom(IFigure figure)
    {
        return figure is AngleMeasurementBase || figure is AngleArc || figure is Slider;
    }

    protected override string Kind
    {
        get
        {
            return "Line";
        }
    }

    /// <summary>"through A at 30°", "through A at angle a"</summary>
    public override string Construction
    {
        get
        {
            if (Dependencies.Count < 2)
            {
                return null;
            }

            return "through " + ConstructionText.Of(Dependencies[0]) + " at " + ConstructionText.AngleValue(AngleSource);
        }
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
    [PropertyGridName("Angle (degrees)")]
    [PropertyGridGroup("Angle")]
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

    /// <summary>The grid's button back from a tied angle (<see cref="Detach"/>), shown while the angle is a figure's</summary>
    [PropertyGridVisible]
    [PropertyGridName("Type the angle")]
    [PropertyGridGroup("Angle")]
    [PropertyGridIcon(PropertyGridIcon.Pencil)]
    public void UntieAngle()
    {
        if (Detach(nameof(Angle)))
        {
            Drawing.RaiseDisplayProperties(this);
        }
    }

    public bool CanEdit(string propertyName)
    {
        switch (propertyName)
        {
            case nameof(Angle):
                return AngleSource is Number;
            case nameof(UntieAngle):
                return this.IsTied(nameof(Angle));
            default:
                return true;
        }
    }

    /// <summary>An angle taken from a figure says which</summary>
    public string Caption(string propertyName, string defaultCaption)
    {
        return propertyName == nameof(Angle) && this.IsTied(propertyName) ? "Angle = " + AngleSource.Name : defaultCaption;
    }

    #region Tied values

    public IEnumerable<string> TiedValueNames
    {
        get { yield return nameof(Angle); }
    }

    public IFigure GetSource(string name)
    {
        return AngleSource;
    }

    public bool Accepts(string name, IFigure figure)
    {
        return CanTakeAngleFrom(figure);
    }

    public bool TieTo(string name, IFigure source)
    {
        return TiedValues.Tie(new IFigure[] { this }, AngleSource, source);
    }

    public bool Detach(string name)
    {
        return this.IsTied(name) && TieTo(name, Number.CreateAuxiliary(Drawing, Angle));
    }

    #endregion
}
