using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// A line through a point at an angle. The angle comes first, so that a line at the angle
/// in the side panel (horizontal until something else is typed) is a single click on a
/// point, or on an empty spot for a new one. For an angle from the drawing instead, click
/// the angle measurement, its arc or a slider first; the line then follows it.
/// </summary>
[Category(BehaviorCategories.Lines)]
[Order(9)]
public class LineAtAngleCreator : FigureCreator
{
    // the next line is likely at the same angle as the last
    static double lastAngle;

    IFigure angleSource;
    AnglePanel panel;

    /// <summary>The side panel: the angle the next line gets, typed, or the figure it comes from</summary>
    [PropertyGridNoUndo]
    public class AnglePanel : IConditionalProperties
    {
        public AnglePanel(LineAtAngleCreator parent)
        {
            this.parent = parent;
        }

        readonly LineAtAngleCreator parent;

        [PropertyGridVisible]
        [PropertyGridFocus]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Angle
        {
            get
            {
                return parent.angleSource is IAngleProvider source ? source.Angle.ToDegrees() : lastAngle;
            }
            set
            {
                lastAngle = value;
            }
        }

        public bool CanEdit(string propertyName)
        {
            return parent.angleSource == null;
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return parent.angleSource == null ? "Angle (degrees)" : "Angle = " + parent.angleSource.Name;
        }

        // the title of the panel
        public override string ToString()
        {
            return "Line at angle";
        }
    }

    public override object PropertyBag
    {
        get
        {
            if (panel == null)
            {
                panel = new AnglePanel(this);
            }

            return panel;
        }
    }

    // before the base raises "construction complete", which is what shows the panel again
    public override void Stopping()
    {
        angleSource = null;
        base.Stopping();
    }

    public override bool IsInInitialState
    {
        get { return angleSource == null && base.IsInInitialState; }
    }

    protected override DependencyList InitExpectedDependencies()
    {
        return DependencyList.Point;
    }

    /// <summary>
    /// An angle measurement, its arc or a slider. Not any other arc, although it has an angle
    /// too: a click on an arc puts the point on it.
    /// </summary>
    static bool CanTakeAngleFrom(IFigure figure)
    {
        return figure is AngleMeasurementBase || figure is AngleArc || figure is Slider;
    }

    protected override IFigure FindFigureInsteadOfPoint(Point unconstrainedCoordinates)
    {
        if (angleSource != null)
        {
            return null;
        }

        var figure = Drawing.Figures.HitTest(unconstrainedCoordinates);
        return figure != null && CanTakeAngleFrom(figure) ? figure : null;
    }

    protected override void Click(Point coordinates)
    {
        var angleFigure = FindFigureInsteadOfPoint(ClickedUnconstrainedCoordinates);
        if (angleFigure == null)
        {
            base.Click(coordinates);
            return;
        }

        StartConstruction();
        angleSource = angleFigure;
        AdvertiseNextDependency();
    }

    protected override IEnumerable<IFigure> CreateFigures()
    {
        var angle = angleSource;
        if (angle == null)
        {
            angle = Number.CreateAuxiliary(Drawing, lastAngle);
            yield return angle;
        }

        yield return LineAtAngle.Create(Drawing, (IPoint)FoundDependencies[0], angle);
    }

    public override string Name
    {
        get { return "Line at Angle"; }
    }

    public override string HintText
    {
        get
        {
            return "Click a point: the line goes through it at the angle in the panel."
                + " To take the angle from an angle or a slider, click that first.";
        }
    }

    public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
    {
        return angleSource != null
            ? "Click a point: the line goes through it at the angle " + angleSource.Name + "."
            : HintText;
    }

    public override FrameworkElement CreateIcon()
    {
        return IconBuilder.BuildIcon()
            .TransparentLine(0.25, 0.75, 1, 0.75, transparency: 0.5)
            .Arc(0.25, 0.75, 0.7, 0.75, 0.601, 0.469)
            .Line(0, 0.95, 1, 0.15)
            .Point(0.25, 0.75)
            .Canvas;
    }
}
