using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// A label attached to something in the plane: it sits at <see cref="Offset"/> from its
/// <see cref="Anchor"/>.
/// </summary>
public abstract class LabelWithOffset : LabelBase, IMovable
{
    /// <summary>
    /// The label's top-left corner from its anchor, in pixels: what the label is attached
    /// to zooms, the text doesn't, and the gap between them shouldn't either (in units of
    /// the plane it shrank onto the point when zoomed out and drifted away when zoomed in).
    /// Files before version 1 stored it in units; see <see cref="UpgradeOffsetFromUnits"/>.
    /// </summary>
    public Point Offset { get; set; }

    /// <summary>
    /// What the offset is measured from, in the plane: the point, the middle of the
    /// segment, the vertex of the angle
    /// </summary>
    public abstract Point Anchor { get; }

    /// <summary>Where the anchor and the offset put the label right now, in the plane</summary>
    protected Point PlaceFromOffset()
    {
        return ToLogical(ToPhysical(Anchor).Plus(Offset));
    }

    /// <summary>Moving the label changes its offset from the anchor</summary>
    public override void MoveToCore(Point newPosition)
    {
        Offset = ToPhysical(newPosition).Minus(ToPhysical(Anchor));
        movedByHand = true;
        base.MoveToCore(newPosition);
    }

    // moved by a drag: drawn where it was put, whatever it covers, until it is next drawn
    bool movedByHand;

    /// <summary>
    /// Whether the label is drawn elsewhere when something covers its place
    /// (<see cref="NameLayout"/>): a point's name, a figure's name, a text attached to a figure
    /// </summary>
    protected virtual bool KeepsClear => false;

    /// <summary>The point the label belongs to, which its other places keep clear of; none but for a point's name or text</summary>
    public virtual IFigure OwnPoint => null;

    /// <summary>The lines and circles that may run through the anchor, which another place of the label must not lie across</summary>
    public virtual IEnumerable<IFigure> FiguresThroughAnchor() => Enumerable.Empty<IFigure>();

    /// <summary>
    /// Where the label is drawn instead of its place, while that is covered: from the anchor
    /// to the label's middle, in pixels; null at its place (<see cref="NameLayout"/>)
    /// </summary>
    public Point? Nudge { get; set; }

    // the size measured last, and for what: measuring lays the text out again, too much to
    // do for every other name at every redraw
    string measuredText;
    double measuredFontSize;
    Size measuredSize;

    /// <summary>The label's size on the canvas, in pixels, measured again only when its text or font changed</summary>
    public Size ShownSize()
    {
        var text = TextBlock.Text ?? "";
        if (text != measuredText || TextBlock.FontSize != measuredFontSize)
        {
            measuredSize = MeasureSize();
            measuredText = text;
            measuredFontSize = TextBlock.FontSize;
        }

        return measuredSize;
    }

    public override void ApplyStyle()
    {
        base.ApplyStyle();

        // (another weight or family is another size)
        measuredText = null;
    }

    /// <summary>Where the label is drawn on the canvas, in pixels; null while it isn't</summary>
    public Rect? ShownBox()
    {
        if (!HasCanvas || !IsShown || !Shape.IsVisible)
        {
            return null;
        }

        var size = ShownSize();
        var topLeft = ToPhysical(Coordinates);
        return size.Width > 0 && size.Height > 0 && topLeft.Exists() ? new Rect(topLeft, size) : null;
    }

    // where the label was drawn last: when it changes, the labels drawn away from their place look again
    Rect? lastShownBox;

    public override void OnRemovingFromDrawing(Drawing drawing)
    {
        NameLayout.Forget(this, drawing);
        base.OnRemovingFromDrawing(drawing);
    }

    /// <summary>The offset is the label's place: undo of a drag puts it back whatever the zoom is by then</summary>
    public override object CapturePlace()
    {
        return Offset;
    }

    public override void RestorePlace(object place)
    {
        Offset = (Point)place;
        if (Drawing != null)
        {
            UpdateVisual();
        }
    }

    /// <summary>
    /// Whether there is an anchor to measure from: something the label is built on, unless a
    /// subclass takes its anchor from elsewhere (a number's caption from the number)
    /// </summary>
    protected virtual bool HasAnchor => !Dependencies.IsEmpty();

    /// <summary>Puts the label where its anchor and offset say; the text is the subclass's business</summary>
    public override void UpdateVisual()
    {
        if (!HasAnchor)
        {
            return;
        }

        Coordinates = PlaceFromOffset();
        if (KeepsClear)
        {
            KeepClear();
        }

        base.UpdateVisual();
        var shown = ShownBox();
        if (shown != lastShownBox)
        {
            lastShownBox = shown;
            NameLayout.Recheck(this);
        }
    }

    /// <summary>
    /// The offset a new label takes, given its default one: that, unless something covers it
    /// or a line through the anchor runs across it, else the nearest free place
    /// (<see cref="NameLayout.FirstPlace"/>)
    /// </summary>
    protected Point FirstOffset(Point offset)
    {
        if (!HasCanvas || Drawing.IsReading)
        {
            return offset;
        }

        var anchor = ToPhysical(Anchor);
        if (!anchor.Exists())
        {
            return offset;
        }

        return NameLayout.FirstPlace(this, anchor.Plus(offset)).Minus(anchor);
    }

    /// <summary>Drawn elsewhere while something covers the label's place (<see cref="NameLayout"/>), unless it was just put there by hand</summary>
    void KeepClear()
    {
        if (movedByHand || !HasCanvas || Drawing.IsReading)
        {
            movedByHand = false;
            NameLayout.Forget(this, Drawing);
            return;
        }

        Coordinates = ToLogical(NameLayout.Place(this, ToPhysical(Coordinates)));
    }

    /// <summary>
    /// A drawing from before version 1 stored the offset in units of the plane; this turns
    /// it into pixels at the zoom the drawing opened at, the best guess for the zoom it was
    /// placed at.
    /// </summary>
    public void UpgradeOffsetFromUnits()
    {
        double unitLength = Drawing.CoordinateSystem.UnitLength;
        Offset = new Point(Offset.X * unitLength, -Offset.Y * unitLength);
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);
        Offset = new Point(element.ReadDouble("OffsetX"), element.ReadDouble("OffsetY"));
    }

    /// <summary>Whether the file says the offset: not for a label that works out its place itself</summary>
    protected virtual bool SavesOffset => true;

    public override void WriteXml(XmlWriter writer)
    {
        base.WriteXml(writer);
        if (SavesOffset)
        {
            writer.WriteAttributeDouble("OffsetX", Offset.X);
            writer.WriteAttributeDouble("OffsetY", Offset.Y);
        }
    }
}

public abstract class Measurement : LabelWithOffset
{
    /// <summary>
    /// Not in the grid: what a measurement says is worked out from what it measures.
    /// Typed there, a text lasted until the next move and was an undo step that undid nothing.
    /// </summary>
    [PropertyGridVisible(false)]
    public override string Text
    {
        get => base.Text; set => base.Text = value;
    }

    /// <summary>
    /// A measurement depends on what it measures, and a dependent figure normally can't be
    /// dragged (its parents move instead). But all that dragging a measurement changes is the
    /// offset of the label from its anchor - so it can, like in the original DG.
    /// </summary>
    public override bool AllowMove()
    {
        return !Locked;
    }
}
