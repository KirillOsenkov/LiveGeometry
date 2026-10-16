using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// A line through a point at an angle. The angle comes first, so that a line at the angle
/// in the side panel (horizontal until something else is typed) is a single click on a
/// point, or on an empty spot for a new one. For an angle from the drawing instead, click
/// the angle measurement, its arc or a slider first; the line then follows it. Right after
/// the line is made its own panel shows the angle (<see cref="TiedValuesPanel"/>), where it
/// can be turned, tied to an angle with a click on it, or typed again.
/// </summary>
[Category(BehaviorCategories.Lines)]
[Order(9)]
public class LineAtAngleCreator : FigureCreator
{
    /// <summary>
    /// The angle the next line gets: the next line is likely at the same angle as the last,
    /// and a line turned in the panel after it sets it too (<see cref="TakeDefaultsFrom"/>).
    /// </summary>
    public static double LastAngle { get; set; }

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

        // not [PropertyGridFocus]: the panel is back after every line, and the angle rarely
        // changes; taking the keyboard each time killed the tool letters and Delete
        [PropertyGridVisible]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Angle
        {
            get => parent.angleSource is IAngleProvider source ? source.Angle.ToDegrees() : LastAngle; set => LastAngle = value;
        }

        /// <summary>
        /// Back to a typed angle before any point is clicked: the clicked figure's angle as
        /// it is now becomes the typed one, and the tool starts over (the panel comes back
        /// editable with it).
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Type the angle")]
        [PropertyGridIcon(PropertyGridIcon.Pencil)]
        public void UntieAngle()
        {
            if (parent.angleSource is IAngleProvider source)
            {
                LastAngle = source.Angle.ToDegrees();
                parent.Restart();
            }
        }

        /// <summary>The next line horizontal: one click on a point after this (<see cref="LineAtAngleCreator.Align"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Horizontal")]
        [PropertyGridIcon(PropertyGridIcon.Horizontal)]
        public void MakeHorizontal()
        {
            parent.Align(axis: 0);
        }

        [PropertyGridVisible]
        [PropertyGridName("Vertical")]
        [PropertyGridIcon(PropertyGridIcon.Vertical)]
        public void MakeVertical()
        {
            parent.Align(axis: 90);
        }

        public bool CanEdit(string propertyName)
        {
            switch (propertyName)
            {
                case nameof(UntieAngle):
                    return parent.angleSource != null;
                // (an angle taken from a figure is let go of for it)
                case nameof(MakeHorizontal):
                    return parent.angleSource != null || !DirectionConstraint.IsAlong(LastAngle, axis: 0);
                case nameof(MakeVertical):
                    return parent.angleSource != null || !DirectionConstraint.IsAlong(LastAngle, axis: 90);
                default:
                    return parent.angleSource == null;
            }
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            if (propertyName != nameof(Angle))
            {
                return defaultCaption;
            }

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

    /// <summary>
    /// The next line's angle along the axis at <paramref name="axis"/> degrees, the nearer way
    /// from the angle the panel shows (0° or 180°, 90° or 270°). An angle clicked in the
    /// drawing is let go of, as "Type the angle" does: the tool starts over, which shows the
    /// panel again; otherwise the panel is shown again here, with its new value and buttons.
    /// </summary>
    void Align(double axis)
    {
        LastAngle = DirectionConstraint.NearestAlong(panel.Angle, axis);
        if (angleSource != null)
        {
            Restart();
        }
        else
        {
            Drawing.RaiseDisplayProperties(panel);
        }
    }

    // before the base raises "construction complete", which is what shows the panel again
    public override void Stopping()
    {
        angleSource = null;
        base.Stopping();
    }

    public override bool IsInInitialState => angleSource == null && base.IsInInitialState;

    protected override DependencyList InitExpectedDependencies()
    {
        return DependencyList.Point;
    }

    /// <summary>What gives an angle under the cursor, when the topmost there does (not under a point)</summary>
    protected override IReadOnlyList<IFigure> FindFiguresInsteadOfPoint(Point unconstrainedCoordinates)
    {
        if (angleSource != null)
        {
            return System.Array.Empty<IFigure>();
        }

        var figures = Drawing.Figures.HitTestAll(unconstrainedCoordinates, f => f.Visible && f.IsHitTestVisible);
        if (figures.Count == 0 || !LineAtAngle.CanTakeAngleFrom(figures[0]))
        {
            return System.Array.Empty<IFigure>();
        }

        return figures.Where(LineAtAngle.CanTakeAngleFrom).ToList();
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

    /// <summary>
    /// A click inside an angle next to its vertex, before anything else, measures the angle
    /// as the Angle tool does and takes the measurement for the line's angle
    /// (<see cref="AngleAtVertex"/>)
    /// </summary>
    protected override bool TakesAngleAtVertex()
    {
        return angleSource == null && FoundDependencies.IsEmpty();
    }

    protected override void ClickAngleAtVertex(AngleAtVertex angle)
    {
        angleSource = MeasureAngleAtVertex(angle);
        AdvertiseNextDependency();
    }

    protected override IEnumerable<IFigure> CreateFigures()
    {
        var angle = angleSource;
        if (angle == null)
        {
            angle = Number.CreateAuxiliary(Drawing, LastAngle);
            yield return angle;
        }

        yield return LineAtAngle.Create(Drawing, (IPoint)FoundDependencies[0], angle);
    }

    /// <summary>A line turned in its panel sets the angle of the next one</summary>
    protected override void TakeDefaultsFrom(ITiedValues created)
    {
        if (created is LineAtAngle line && line.AngleSource is Number number)
        {
            LastAngle = number.Value;
        }
    }

    protected override string CreatedFigureHint(ITiedValues values)
    {
        return values.IsTied(nameof(LineAtAngle.Angle))
            ? "click another angle or a slider to take the angle from it instead."
            : "set its angle in the panel, or click an angle or a slider to take the angle from it.";
    }

    public override string Name => "Line at Angle";

    public override string HintText
    {
        get
        {
            return "Click a point: the line goes through it at the angle in the panel."
                + " To take the angle from an angle or a slider, click that first, or click inside an angle next to its vertex.";
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
            .AccentLine(0, 0.95, 1, 0.15)
            .Point(0.25, 0.75)
            .Canvas;
    }
}
