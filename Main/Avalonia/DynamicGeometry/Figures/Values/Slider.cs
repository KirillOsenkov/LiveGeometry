using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;

namespace DynamicGeometry;

/// <summary>
/// An adjustable number with a handle: a horizontal track from an anchor to a knob, captioned
/// "a = 2.00". The value is the length of the track in units, so a slider is a length (the
/// radius for Circle by Radius, the distance for Translation), an angle in degrees (Rotation)
/// and a Number an expression names, all at once. Dragging the knob changes the value,
/// dragging anything else moves the slider. The parts are figures of the library - a free
/// point, a translated point kept on the horizontal through it, a segment, a label - held
/// inside this one: the drawing sees a single figure and the file a single element, and the
/// parts never have to be resolved by name.
/// </summary>
public class Slider : CompositeFigure, INumber, ILengthProvider, IAngleProvider, IMovableParts
{
    // the knob's direction: an angle provider saying 0 degrees; not a child, it has no shape
    readonly Number horizontal = new Number() { Value = 0 };

    readonly WholeHandle wholeHandle;

    public Slider()
    {
        // the parts are looked up by name only through the composite (Figures[name] looks
        // inside), so their names must not be ones an expression could say
        Anchor = new FreePoint() { Name = "slider anchor" };
        Knob = new SliderKnob(this) { Name = "slider knob" };
        Knob.SetSources(Anchor, distanceSource: null, directionSource: horizontal);
        Track = new Segment() { Name = "slider track", Dependencies = new IFigure[] { Anchor, Knob } };
        Caption = new SliderCaption(this) { Name = "slider caption", Dependencies = new IFigure[] { Anchor, Knob } };
        AddChild(Anchor);
        AddChild(Knob);
        AddChild(Track);
        AddChild(Caption);
        wholeHandle = new WholeHandle(this);

        // of the figures under the cursor the topmost wins: the knob over a line crossing it
        Layer = ZOrder.Points;
    }

    /// <summary>Where the slider sits; a free point, yellow</summary>
    public FreePoint Anchor { get; }

    /// <summary>What the user drags to change the value; slides along the horizontal, green</summary>
    public SliderKnob Knob { get; }

    /// <summary>From the anchor to the knob: its length is the value</summary>
    public Segment Track { get; }

    /// <summary>"a = 2.00", above the anchor</summary>
    public SliderCaption Caption { get; }

    #region Value and place

    double minimum;
    double maximum = double.PositiveInfinity;

    /// <summary>
    /// The number the slider stands for: <see cref="Minimum"/> at the anchor, one more per
    /// unit of track, up to <see cref="Maximum"/>. Set it and the knob moves; anything built
    /// on the slider follows.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridPreferredEditor("UpDown")]
    public double Value
    {
        get
        {
            return minimum + Knob.Distance;
        }
        set
        {
            var origin = Anchor.Coordinates;
            double clamped = System.Math.Max(minimum, System.Math.Min(maximum, value));
            Knob.MoveToCore(new Point(origin.X + clamped - minimum, origin.Y));
            OnChanged();
        }
    }

    /// <summary>The value at the anchor, 0 unless a file or the grid says otherwise</summary>
    [PropertyGridVisible]
    [PropertyGridGroup("Range")]
    [PropertyGridPreferredEditor("UpDown")]
    public double Minimum
    {
        get
        {
            return minimum;
        }
        set
        {
            if (!value.IsValidValue() || value > maximum)
            {
                return;
            }

            double kept = Value;
            minimum = value;
            Value = kept;
        }
    }

    /// <summary>Where the knob stops; no stop unless a file or the grid says otherwise</summary>
    [PropertyGridVisible]
    [PropertyGridGroup("Range")]
    [PropertyGridPreferredEditor("UpDown")]
    public double Maximum
    {
        get
        {
            return maximum;
        }
        set
        {
            if (double.IsNaN(value) || value < minimum)
            {
                return;
            }

            double kept = Value;
            maximum = value;
            Value = kept;
        }
    }

    /// <summary>How far the knob may go from the anchor</summary>
    public double Span
    {
        get { return maximum - minimum; }
    }

    /// <summary>Where the anchor is; the knob keeps its distance</summary>
    public Point Position
    {
        get
        {
            return Anchor.Coordinates;
        }
        set
        {
            Anchor.MoveToCore(value);
            OnChanged();
        }
    }

    /// <summary>The parts follow, then whatever is built on the slider</summary>
    void OnChanged()
    {
        RaisePropertyChanged("Value");
        if (Drawing != null)
        {
            // the descendants of a figure include the figure itself
            this.RecalculateAllDependents();
        }
    }

    [PropertyGridVisible]
    [Domain(0, 10)]
    public int Decimals
    {
        get
        {
            return Caption.DecimalsToShow;
        }
        set
        {
            Caption.DecimalsToShow = value;
        }
    }

    public double Length
    {
        get { return Value; }
    }

    /// <summary>Angle providers speak radians; the slider itself is in degrees, like a Number</summary>
    public double Angle
    {
        get { return Value.ToRadians(); }
    }

    public override Point Center
    {
        get { return Track.Center; }
    }

    #endregion

    #region Name

    // a, b, c: what an expression says, next to points A, B, C (the letters lines and circles
    // are named with too, from g and c on)
    public override string GenerateFigureName(List<string> blacklist)
    {
        return GenerateLetterName(this, "a", blacklist);
    }

    /// <summary>A composite dumps its parts by default</summary>
    public override string ToString()
    {
        return Name;
    }

    /// <summary>A slider goes by its name, as a number does: "radius a"</summary>
    public override string Noun
    {
        get
        {
            return null;
        }
    }

    /// <summary>"= 2"</summary>
    public override string Construction
    {
        get
        {
            return "= " + ConstructionText.Number(Value);
        }
    }

    /// <summary>The property grid's title: "Slider a"</summary>
    protected override string Kind
    {
        get
        {
            return "Slider";
        }
    }

    /// <summary>The caption shows the name</summary>
    [PropertyGridVisible]
    [PropertyGridDisallowMultiEdit]
    public override string Name
    {
        get
        {
            return base.Name;
        }
        set
        {
            base.Name = value;
            if (Drawing != null)
            {
                Caption.UpdateVisual();
            }
        }
    }

    #endregion

    #region Style

    /// <summary>The track's line style stands for the slider's; the points and the caption keep the styles of their kinds</summary>
    public override IFigureStyle Style
    {
        get
        {
            return Track.Style;
        }
        set
        {
            Track.Style = value;
        }
    }

    public override void OnAddingToCanvas(Canvas newContainer)
    {
        // before the parts take the styles of their kinds: the track is a segment and would
        // take the default line, which the slider would then pass for its own
        EnsureStyleAssigned();
        base.OnAddingToCanvas(newContainer);
    }

    #endregion

    #region Hit testing and dragging

    /// <summary>
    /// A hit on any part is a hit on the slider - the parts are never handed out, a tool
    /// that took the knob for a point would depend on something not in the drawing.
    /// </summary>
    public override IFigure HitTest(Point point, System.Predicate<IFigure> filter)
    {
        return FindPart(point) != null && filter(this) ? this : null;
    }

    /// <summary>The visible part under the point; the knob first, it sits on the track and on the anchor at zero</summary>
    IFigure FindPart(Point point)
    {
        foreach (var part in new IFigure[] { Knob, Anchor, Caption, Track })
        {
            if (part.Visible && part.HitTest(point) != null)
            {
                return part;
            }
        }

        return null;
    }

    /// <summary>The knob changes the value, the anchor and everything else move the slider</summary>
    public IMovable FindMovablePart(Point point)
    {
        var part = FindPart(point);
        if (part == Knob)
        {
            return Knob;
        }

        if (part == Anchor)
        {
            return Anchor;
        }

        return wholeHandle;
    }

    public IMovable WholePart
    {
        get { return wholeHandle; }
    }

    /// <summary>
    /// Dragging the track or the caption: the anchor moves by as much as the cursor, without
    /// jumping under it the way a dragged point does.
    /// </summary>
    class WholeHandle : IMovable
    {
        readonly Slider slider;

        public WholeHandle(Slider slider)
        {
            this.slider = slider;
        }

        public Point Coordinates
        {
            get { return slider.Anchor.Coordinates; }
        }

        public bool AllowMove()
        {
            return !slider.Locked;
        }

        public void MoveTo(Point position)
        {
            slider.Anchor.MoveTo(position);
        }
    }

    #endregion

    #region Parts

    /// <summary>The knob: a point sliding along the horizontal through the anchor, never to its left, nor past the slider's span</summary>
    public class SliderKnob : TranslatedPoint
    {
        readonly Slider slider;

        public SliderKnob(Slider slider)
        {
            this.slider = slider;
        }

        public override void MoveToCore(Point newPosition)
        {
            var source = Source;
            if (source != null)
            {
                double left = source.Coordinates.X;
                double right = left + slider.Span;
                if (newPosition.X < left)
                {
                    newPosition = new Point(left, newPosition.Y);
                }
                else if (newPosition.X > right)
                {
                    newPosition = new Point(right, newPosition.Y);
                }
            }

            base.MoveToCore(newPosition);
        }
    }

    /// <summary>
    /// "a = 2.00" just above the anchor, left-aligned with it, a fixed few pixels away so that
    /// it keeps its place at every zoom. Not draggable on its own: a drag moves the slider.
    /// </summary>
    public class SliderCaption : LabelWithOffset
    {
        // pixels between the points and the text's line box, which has air of its own under the letters
        const double gap = 1;

        readonly Slider slider;

        public SliderCaption(Slider slider)
        {
            this.slider = slider;
        }

        public override Point Anchor
        {
            get { return slider.Anchor.Coordinates; }
        }

        public override bool AllowMove()
        {
            return false;
        }

        public override void UpdateVisual()
        {
            if (Drawing == null)
            {
                return;
            }

            Text = NameDisplay.Format(slider.Name) + " = " + Math.Round(slider.Value, DecimalsToShow).ToString();
            var size = MeasureSize();

            // left-aligned with the anchor, and above the bigger of the two points: a knob
            // bigger than the anchor (a drawing's own styles) sat on the text near the start
            double anchorRadius = slider.Anchor.Shape.Width / 2;
            double knobWidth = slider.Knob.Shape.Width;
            double pointRadius = knobWidth.IsValidValue() ? System.Math.Max(anchorRadius, knobWidth / 2) : anchorRadius;
            Offset = new Point(-anchorRadius, -(size.Height + pointRadius + gap));
            base.UpdateVisual();
        }
    }

    #endregion

    #region Serialization

    public override void WriteXml(XmlWriter writer)
    {
        base.WriteXml(writer);
        var position = Position;
        writer.WriteAttributeDouble("X", position.X);
        writer.WriteAttributeDouble("Y", position.Y);
        writer.WriteAttributeDouble("Value", Value);
        if (minimum != 0)
        {
            writer.WriteAttributeDouble("Minimum", minimum);
        }

        if (!double.IsPositiveInfinity(maximum))
        {
            writer.WriteAttributeDouble("Maximum", maximum);
        }

        if (Decimals != Settings.DisplayDecimals)
        {
            writer.WriteAttributeDouble("Decimals", Decimals);
        }
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);
        if (element.Attribute("Decimals") != null)
        {
            Decimals = (int)element.ReadDouble("Decimals");
        }

        // the range before the value, which is clamped to it
        if (element.Attribute("Minimum") != null)
        {
            minimum = element.ReadDouble("Minimum");
        }

        if (element.Attribute("Maximum") != null)
        {
            maximum = element.ReadDouble("Maximum");
        }

        Anchor.MoveToCore(new Point(element.ReadDouble("X"), element.ReadDouble("Y")));
        Value = element.ReadDouble("Value");
    }

    #endregion
}
