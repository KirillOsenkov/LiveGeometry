using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// A text attached to a figure (2026-10-17): written next to a point, a segment, a line, a
/// ray or a circle as its name would be, or in the middle of a polygon
/// (<see cref="FigureAnchor"/>), and moving with it,
/// but saying whatever is typed - "r", "hypotenuse", "x = 4√2" - so that a name doesn't have
/// to be made into a comment. The Text tool makes one with a click on a figure; it is a
/// figure of its own (the Figure List has it, Delete takes it) and goes when its figure
/// goes. The text is plain: a free label's [...] parts make it depend on the figures they
/// name, and this one depends on its figure alone. It keeps clear as a name does
/// (<see cref="NameLayout"/>). The player has the twin (attachedLabel.js).
/// </summary>
public class AttachedLabel : LabelWithOffset
{
    bool placed;

    /// <summary>The figure the text is attached to</summary>
    public IFigure Figure => Dependencies.Count > 0 ? Dependencies[0] : null;

    /// <summary>"on segment AB": "Text “r” on segment AB"</summary>
    public override string Construction => Figure != null ? "on " + ConstructionText.Of(Figure) : null;

    /// <summary>Plain text, not the free label's box that puts references to figures in [...]: they would not be worked out here</summary>
    [PropertyGridVisible]
    [PropertyGridMultiline]
    [PropertyGridDisallowMultiEdit]
    public override string Text
    {
        get => base.Text; set => base.Text = value;
    }

    /// <summary>No numbers in a plain text: no Decimals</summary>
    public override bool ShowsDecimals()
    {
        return false;
    }

    public override Point Anchor => FigureAnchor.Of(Figure, Drawing?.CoordinateSystem);

    protected override bool KeepsClear => true;

    public override IFigure OwnPoint => Figure as IPoint;

    /// <summary>A point's lines and circles; a segment, a line or a circle itself</summary>
    public override IEnumerable<IFigure> FiguresThroughAnchor()
    {
        var figure = Figure;
        if (figure == null)
        {
            return Enumerable.Empty<IFigure>();
        }

        return figure is IPoint ? figure.Dependents.Concat(figure.Dependencies) : new[] { figure };
    }

    /// <summary>Puts the text with its top-left corner here, logical: from now on its place</summary>
    public void PlaceAt(Point topLeft)
    {
        placed = true;
        Offset = ToPhysical(topLeft).Minus(ToPhysical(Anchor));
    }

    /// <summary>Lets go of the figure: a free text where this one is drawn (<see cref="TextAttachment.Detach"/>)</summary>
    [PropertyGridVisible]
    [PropertyGridName("Detach")]
    [PropertyGridIcon(PropertyGridIcon.Unlock)]
    public void Detach()
    {
        var free = TextAttachment.Detach(this);
        free.Drawing.RaiseDisplayProperties(free);
    }

    /// <summary>A text attached to a figure is dragged by itself: that changes its offset from the anchor, not the figure</summary>
    public override bool AllowMove()
    {
        return !Locked;
    }

    protected override ZOrder DefaultLayer()
    {
        return ZOrder.PointLabels;
    }

    public override void UpdateVisual()
    {
        var figure = Figure;
        if (figure == null || Drawing == null)
        {
            return;
        }

        if (!placed && !Text.IsEmpty())
        {
            // placed for good once the drawing is there to keep clear of (a file written by
            // hand that says no place is still being read)
            Offset = FigureAnchor.DefaultOffset(figure, MeasureSize(), Drawing.CoordinateSystem);
            if (HasCanvas && !Drawing.IsReading)
            {
                placed = true;
                Offset = FirstOffset(Offset);
            }
        }

        base.UpdateVisual();
        bool shown = Visible && figure.Visible && figure.Exists;
        if (figure is LineBase && figure is not Segment)
        {
            // a line that runs outside the window has nowhere to write it
            shown = shown && FigureAnchor.IsOnScreen(ToPhysical(Anchor), Drawing.CoordinateSystem);
            if (shown)
            {
                var topLeft = ToPhysical(Coordinates);
                var kept = FigureAnchor.KeptOnScreen(topLeft, MeasureSize(), Drawing.CoordinateSystem);
                if (kept != topLeft && kept.Exists())
                {
                    Coordinates = ToLogical(kept);
                    Shape.MoveTo(kept);
                }
            }
        }

        // a hidden figure keeps its text to itself
        Shape.IsVisible = shown;
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);
        Text = Label.Unescape(element.ReadString("Text") ?? "");

        // a file written by hand may leave the place out: then the text goes where a new one would
        placed = element.Attribute("OffsetX") != null || element.Attribute("OffsetY") != null;
    }

    public override void WriteXml(XmlWriter writer)
    {
        base.WriteXml(writer);
        writer.WriteAttributeString("Text", Label.Escape(Text));
    }
}
