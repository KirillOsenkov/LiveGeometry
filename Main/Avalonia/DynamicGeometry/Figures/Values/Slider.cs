using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using Avalonia;

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
        Knob = new SliderKnob() { Name = "slider knob" };
        Knob.SetSources(Anchor, distanceSource: null, directionSource: horizontal);
        Track = new Segment() { Name = "slider track", Dependencies = new IFigure[] { Anchor, Knob } };
        Caption = new SliderCaption(this) { Name = "slider caption", Dependencies = new IFigure[] { Anchor, Knob } };
        AddChild(Anchor);
        AddChild(Knob);
        AddChild(Track);
        AddChild(Caption);
        wholeHandle = new WholeHandle(this);

        // of the figures under the cursor the topmost wins: the knob over a line crossing it
        ZIndex = (int)ZOrder.Points;
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

    /// <summary>
    /// The number the slider stands for, never negative. Set it and the knob moves; anything
    /// built on the slider follows.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridPreferredEditor("UpDown")]
    public double Value
    {
        get
        {
            return Knob.Distance;
        }
        set
        {
            var origin = Anchor.Coordinates;
            Knob.MoveToCore(new Point(origin.X + value, origin.Y));
            OnChanged();
        }
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

    // a, b, c: what an expression says, next to points A, B, C. Not e (the constant), not x
    // and y (the axes and the variable of a function), not the letters that read as digits.
    const string alphabet = "abcdfghkmnpqrstuvwz";

    public override string GenerateFigureName(List<string> blacklist)
    {
        for (int i = 0; ; i++)
        {
            foreach (var letter in alphabet)
            {
                var candidate = i == 0 ? letter.ToString() : letter + i.ToString();
                if (this.NameAvailable(candidate) && (blacklist == null || !blacklist.Contains(candidate)))
                {
                    return candidate;
                }
            }
        }
    }

    /// <summary>The property grid's title; a composite dumps its parts by default</summary>
    public override string ToString()
    {
        return Name;
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

    /// <summary>The knob: a point sliding along the horizontal through the anchor, never to its left</summary>
    public class SliderKnob : TranslatedPoint
    {
        public override void MoveToCore(Point newPosition)
        {
            var source = Source;
            if (source != null && newPosition.X < source.Coordinates.X)
            {
                newPosition = new Point(source.Coordinates.X, newPosition.Y);
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

            Text = slider.Name + " = " + Math.Round(slider.Value, DecimalsToShow).ToString();
            var size = MeasureSize();
            double pointRadius = slider.Anchor.Shape.Width / 2;
            Offset = new Point(-pointRadius, -(size.Height + pointRadius + Math.CursorTolerance));
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

        Anchor.MoveToCore(new Point(element.ReadDouble("X"), element.ReadDouble("Y")));
        Value = element.ReadDouble("Value");
    }

    #endregion
}
