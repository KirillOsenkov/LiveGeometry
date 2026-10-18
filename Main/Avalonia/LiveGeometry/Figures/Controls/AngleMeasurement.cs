using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace LiveGeometry;

public class AngleMeasurementBase : Measurement, IAngleProvider
{
    private bool radians;

    [PropertyGridVisible]
    [PropertyGridSection(PropertyGridSection.Display, 10)]
    public bool Radians
    {
        get => radians;
        set
        {
            radians = value;
            UpdateVisual();
        }
    }

    protected override string Kind => "Angle";

    /// <summary>"ABC": "Angle ABC"</summary>
    public override string Construction => Dependencies.Count < 3 ? null : ConstructionText.Points(Dependencies[1], Dependencies[0], Dependencies[2]);

    /// <summary>The number shown: the angle in the unit chosen</summary>
    public virtual double Measure
    {
        get
        {
            var angle = Angle;
            return (Radians) ? angle : angle.ToDegrees();
        }
    }

    /// <summary>The angle in radians, 0 to 2π: counterclockwise from the first side to the second</summary>
    public virtual double Angle => Math.OAngle(Point(1), Point(0), Point(2));

    public override Point Anchor => Point(0);

    public override void UpdateExistence()
    {
        base.UpdateExistence();
        if (Exists && !AngleArc.HasSides(this))
        {
            Exists = false;
        }
    }

    public override void UpdateVisual()
    {
        // the text first: where a number that places itself goes depends on its size
        var text = Math.Round(Measure, DecimalsToShow).ToString();
        Text = (Radians) ? text + " rad" : text + "°";
        Arrange();
        base.UpdateVisual();
    }

    /// <summary>Works out where the number goes before it is put there, when it places itself</summary>
    protected virtual void Arrange()
    {
    }

    public override void ReadXml(System.Xml.Linq.XElement element)
    {
        base.ReadXml(element);
        radians = element.ReadBool("Radians", false);
    }

    public override void WriteXml(System.Xml.XmlWriter writer)
    {
        base.WriteXml(writer);
        if (Radians)
        {
            writer.WriteAttributeBool("Radians", true);
        }
    }
}

public class AngleMeasurement : AngleMeasurementBase, IConditionalProperties, IHasSweep
{
    AngleSweep sweep = DefaultSweep;

    /// <summary>A new angle is the one under 180°, whichever way round its sides were clicked</summary>
    public const AngleSweep DefaultSweep = AngleSweep.Smaller;

    /// <summary>
    /// Which of the two angles at the vertex the number says (<see cref="AngleSweep"/>);
    /// the mark next to it, its companion, shows the same one. What is built on the
    /// angle (a bisector, a rotation by it) follows.
    /// </summary>
    [PropertyGridVisible]
    public AngleSweep Sweep
    {
        get => sweep;
        set
        {
            if (sweep != value)
            {
                sweep = value;
                if (Drawing != null)
                {
                    this.RecalculateAllDependents();
                }
            }

            AngleArc.SyncCompanionSweep(this, value);
        }
    }

    /// <summary>The measure of the angle the sweep chooses</summary>
    public override double Angle => Sweep.Measure(base.Angle);

    /// <summary>Without its arc (deleted) the number has no arcs to count</summary>
    public bool CanEdit(string propertyName)
    {
        return propertyName != nameof(ArcCount) || FindArc() != null;
    }

    public string Caption(string propertyName, string defaultCaption)
    {
        return defaultCaption;
    }

    /// <summary>
    /// The arc that was created together with this label: same vertex, same two sides.
    /// Null if it has been deleted.
    /// </summary>
    AngleArc FindArc()
    {
        return AngleArc.FindCompanion(this) as AngleArc;
    }

    /// <summary>
    /// The arcs belong to the <see cref="AngleArc"/>, a figure of its own. They can be set
    /// from the label too: with no arcs there is nothing left of the arc to click on.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridName("Arcs")]
    [PropertyGridSection(PropertyGridSection.Display, 30)]
    [Domain(0, 3)]
    [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
    public int ArcCount
    {
        get
        {
            var arc = FindArc();
            return arc != null ? arc.ArcCount : 0;
        }
        set
        {
            var arc = FindArc();
            if (arc != null)
            {
                arc.ArcCount = value;
            }
        }
    }

    #region Place

    bool placedByHand;

    /// <summary>
    /// Whether the number goes where <see cref="AngleLabelLayout"/> puts it: inside its
    /// angle, past the mark, apart from the other numbers at the vertex. A new number does;
    /// one dragged stays where it was dragged to (relative to the vertex, as before), and
    /// the others keep clear of it. A file says OffsetX and OffsetY only for a number placed
    /// by hand; files from before this say them for every number, which stay where they were.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridName("Place automatically")]
    [PropertyGridSection(PropertyGridSection.Display, 40)]
    [PropertyGridCustomValueProvider(typeof(PlaceValue))]
    public bool PlacesItself
    {
        get => !placedByHand;
        set
        {
            if (value == !placedByHand)
            {
                return;
            }

            // unticked, the number stays where it was put last; ticked, it is placed afresh
            placedByHand = !value;
            LayoutPlacement = null;
            if (HasCanvas)
            {
                UpdateVisual();
            }
        }
    }

    /// <summary>The row as the grid sets it: undo puts back the offset a number placed by hand had</summary>
    public class PlaceValue : PropertyValue, IRestorableValue
    {
        public object CaptureState()
        {
            return ((AngleMeasurement)Parent).CapturePlace();
        }

        public void RestoreState(object state)
        {
            ((AngleMeasurement)Parent).RestorePlace(state);
        }
    }

    record SavedPlace(bool PlacedByHand, Point Offset);

    /// <summary>Dragged, the number stays where it is put</summary>
    public override void MoveToCore(Point newPosition)
    {
        bool wasPlacing = PlacesItself;
        placedByHand = true;
        base.MoveToCore(newPosition);
        if (wasPlacing)
        {
            RaisePropertyChanged(nameof(PlacesItself));
        }
    }

    /// <summary>Undo of a drag puts back a number that placed itself, as well as the offset</summary>
    public override object CapturePlace()
    {
        return new SavedPlace(placedByHand, Offset);
    }

    public override void RestorePlace(object place)
    {
        if (place is SavedPlace saved)
        {
            bool changed = saved.PlacedByHand != placedByHand;
            placedByHand = saved.PlacedByHand;
            LayoutPlacement = null;
            base.RestorePlace(saved.Offset);
            if (changed)
            {
                RaisePropertyChanged(nameof(PlacesItself));
            }

            return;
        }

        base.RestorePlace(place);
    }

    /// <summary>A number that places itself saves no offset: read back, it places itself again</summary>
    protected override bool SavesOffset => placedByHand;

    bool isLeaving;

    /// <summary>Whether the number is one the layout at the vertex keeps clear of or places: shown, of an angle there</summary>
    public bool TakesPartInLayout(IFigure vertex)
    {
        return !isLeaving
            && Visible
            && Exists
            && Drawing != null
            && Dependencies.Count == 3
            && Dependencies[0] == vertex
            && AngleArc.HasSides(this);
    }

    protected override void Arrange()
    {
        if (Visible && Exists && HasCanvas && Drawing.CoordinateSystem.PhysicalSize.X > 0 && AngleArc.HasSides(this))
        {
            AngleLabelLayout.ArrangeAt(this);
        }
    }

    /// <summary>Puts the shape where the offset says, without working the layout out again (another number at the vertex just did)</summary>
    public void MoveToOffset()
    {
        if (HasCanvas)
        {
            Coordinates = PlaceFromOffset();
            Shape.MoveTo(ToPhysical(Coordinates));
        }
    }

    /// <summary>The angle and the mark, as the layout takes them</summary>
    public AngleLabelLayout.Wedge LayoutWedge()
    {
        var vertex = Point(0);
        var counterclockwise = Math.OAngle(Point(1), vertex, Point(2));
        var measure = Sweep.Measure(counterclockwise);
        return new AngleLabelLayout.Wedge()
        {
            Vertex = ToPhysical(vertex),
            Start = Math.GetAngle(vertex, Point(1)),
            Sweep = measure,
            IsClockwise = Sweep.IsClockwise(counterclockwise),
            Clearance = MarkReach(measure),
            Size = LayoutSize(),
            Previous = LayoutPlacement
        };
    }

    /// <summary>
    /// Where the layout put the number last, relative to its angle: it starts from there next
    /// time, so that it moves smoothly with the figures (not saved: read back, a number is placed afresh)
    /// </summary>
    public AngleLabelLayout.Placement? LayoutPlacement { get; set; }

    /// <summary>In pixels: how far from the vertex the mark reaches, its outer arc or the far corner of the square of a right angle</summary>
    double MarkReach(double measure)
    {
        var arc = FindArc();
        if (arc == null || !arc.Visible || !arc.Exists || arc.ArcCount == 0)
        {
            return AngleLabelLayout.NoMarkClearance;
        }

        double stroke = arc.Shape.StrokeThickness;
        if (System.Math.Abs(measure - Math.PI / 2) < AngleArc.RightAngleTolerance)
        {
            return arc.Size * System.Math.Sqrt(2) + stroke / 2;
        }

        return arc.Size + (arc.ArcCount - 1) * (stroke + AngleArc.ArcSpacing) + stroke / 2;
    }

    string measuredText;
    Size measuredSize;

    /// <summary>The size of the box, measured again only when the text shown changes (or the style: <see cref="ApplyStyle"/>)</summary>
    public Size LayoutSize()
    {
        var shown = TextBlock.Text ?? "";
        if (shown != measuredText)
        {
            measuredSize = MeasureSize();
            measuredText = shown;
        }

        return measuredSize;
    }

    public override void ApplyStyle()
    {
        measuredText = null;
        base.ApplyStyle();
    }

    /// <summary>The numbers left at the vertex place themselves again without this one</summary>
    public override void OnRemovingFromDrawing(Drawing drawing)
    {
        base.OnRemovingFromDrawing(drawing);
        if (Dependencies.Count != 3)
        {
            return;
        }

        var vertex = Dependencies[0];
        var other = vertex.Dependents
            .OfType<AngleMeasurement>()
            .FirstOrDefault(number => number != this && number.PlacesItself && number.TakesPartInLayout(vertex));
        if (other != null)
        {
            isLeaving = true;
            try
            {
                other.UpdateVisual();
            }
            finally
            {
                isLeaving = false;
            }
        }
    }

    #endregion

    public override void ReadXml(System.Xml.Linq.XElement element)
    {
        base.ReadXml(element);
        sweep = element.ReadSweep(DefaultSweep);

        // (a file says where a number placed by hand goes)
        placedByHand = element.Attribute("OffsetX") != null;
    }

    public override void WriteXml(System.Xml.XmlWriter writer)
    {
        base.WriteXml(writer);
        if (sweep != DefaultSweep)
        {
            writer.WriteAttributeString("Sweep", sweep.ToString());
        }
    }
}

public class HorizontalAngleMeasurement : AngleMeasurementBase
{
    /// <summary>"of AB to the x axis"</summary>
    public override string Construction => Dependencies.Count < 2 ? null : "of " + ConstructionText.Points(Dependencies[0], Dependencies[1]) + " to the x axis";

    /// <summary>The angle of the segment to the x axis, counterclockwise, in radians (the base reads a third point)</summary>
    public override double Angle => Math.OHAngle(Point(0), Point(1));

    public override double Measure => Radians ? Angle : Angle.ToDegrees();

    public override void UpdateVisual()
    {
        if (Dependencies.IsEmpty())
        {
            return;
        }

        // placed by its top left corner as every label is (ControlBase.UpdateVisual):
        // it centered the shape, which has no Width to halve, and the shape stayed in the
        // top left corner of the window
        Coordinates = PlaceFromOffset();
        Shape.MoveTo(ToPhysical(Coordinates));
        Text = Angle
            .ToDegrees()
            .ToDegreeString();
    }
}
