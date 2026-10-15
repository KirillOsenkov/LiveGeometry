using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry;

public class Ray : LineBase, ILine, IFixableDirection
{
    /// <summary>
    /// The way from the start through the second point, in degrees: typed, fixed and freed
    /// as a segment's (<see cref="Segment.Direction"/>, <see cref="DirectionConstraint"/>).
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridGroup("Direction")]
    [PropertyGridPreferredEditor("UpDown")]
    [PropertyGridCustomValueProvider(typeof(DirectionPropertyValue))]
    public virtual double Direction
    {
        get
        {
            return DirectionConstraint.Of(Coordinates);
        }
        set
        {
            DirectionConstraint.SetDirection(this, value);
        }
    }

    public IPoint DirectionEndpoint => DirectionConstraint.Endpoint(this);

    public virtual bool CanEdit(string propertyName)
    {
        return !DirectionConstraint.IsMember(propertyName) || DirectionConstraint.CanEdit(this, propertyName);
    }

    public virtual string Caption(string propertyName, string defaultCaption)
    {
        return defaultCaption;
    }

    public override PointPair OnScreenCoordinates
    {
        get
        {
            var c = Coordinates;
            c = Math.GetLineFromSegment(c, CanvasLogicalBorders);
            c.P1 = Coordinates.P1;
            return c;
        }
    }

    public override double GetNearestParameterFromPoint(Avalonia.Point point)
    {
        var parameter = base.GetNearestParameterFromPoint(point);
        if (parameter < 0)
        {
            parameter = 0;
        }
        return parameter;
    }

    public override IFigure HitTest(Avalonia.Point point)
    {
        var hit = base.HitTest(point) != null;
        var line = Coordinates;

        // the start is on the ray however the rounding of a point there went
        var inside = Math.GetProjection(point, line).Ratio >= -Math.EndTolerance(line);
        if (hit && inside)
        {
            return this;
        }
        return null;
    }

    public override Tuple<double, double> GetParameterDomain()
    {
        return new Tuple<double, double>(0, base.GetParameterDomain().Item2);
    }

    // A first: ray AB starts at A and isn't ray BA
    protected override IReadOnlyList<string> NamesFromDependencies()
    {
        return NamesFromPoints(PointOrder.Fixed);
    }

    protected override string Kind
    {
        get
        {
            return "Ray";
        }
    }

    public override string Noun
    {
        get
        {
            return "ray";
        }
    }

    [PropertyGridVisible]
    [PropertyGridName("Convert to line")]
    [PropertyGridIcon(PropertyGridIcon.Line)]
    public void ConvertToLine()
    {
        LineTwoPoints.Convert(this, Factory.CreateLineTwoPoints(this.Drawing, this.Dependencies));
    }

    [PropertyGridVisible]
    [PropertyGridName("Convert to segment")]
    [PropertyGridIcon(PropertyGridIcon.Segment)]
    public void ConvertToSegment()
    {
        LineTwoPoints.Convert(this, Factory.CreateSegment(this.Drawing, this.Dependencies));
    }

    [PropertyGridVisible]
    [PropertyGridName("Reverse")]
    [PropertyGridIcon(PropertyGridIcon.Reverse)]
    public void Reverse()
    {
        LineTwoPoints.Convert(this, Factory.CreateRay(this.Drawing, this.Dependencies.Reverse().ToList()));
    }

    [PropertyGridVisible]
    [PropertyGridName("Fix direction")]
    [PropertyGridGroup("Direction")]
    [PropertyGridIcon(PropertyGridIcon.Lock)]
    public void FixDirection()
    {
        DirectionConstraint.FixDirection(this);
        Drawing.RaiseDisplayProperties(this);
    }

    [PropertyGridVisible]
    [PropertyGridName("Free direction")]
    [PropertyGridGroup("Direction")]
    [PropertyGridIcon(PropertyGridIcon.Unlock)]
    public void FreeDirection()
    {
        DirectionConstraint.FreeDirection(this);
        Drawing.RaiseDisplayProperties(this);
    }

    [PropertyGridVisible]
    [PropertyGridName("Horizontal")]
    [PropertyGridGroup("Direction")]
    [PropertyGridIcon(PropertyGridIcon.Horizontal)]
    public void MakeHorizontal()
    {
        DirectionConstraint.Align(this, axis: 0);
        Drawing.RaiseDisplayProperties(this);
    }

    [PropertyGridVisible]
    [PropertyGridName("Vertical")]
    [PropertyGridGroup("Direction")]
    [PropertyGridIcon(PropertyGridIcon.Vertical)]
    public void MakeVertical()
    {
        DirectionConstraint.Align(this, axis: 90);
        Drawing.RaiseDisplayProperties(this);
    }
}