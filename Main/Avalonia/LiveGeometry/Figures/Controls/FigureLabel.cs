using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// The name of a line, ray, segment or circle, written next to it (as a <see cref="PointLabel"/>
/// is a point's name): a label that depends on the figure and sits <see cref="LabelWithOffset.Offset"/>
/// pixels from an anchor on it - the middle of a segment, a point on the upper left of a
/// circle, and for a line or a ray a point of its visible part near the edge of the window,
/// so that the name is always on screen however far the line runs. Shown and hidden through
/// the figure's "Show name" (<see cref="FigureBase.HasNameLabel"/>), which adds and removes
/// it; not in the Figure List, since it is part of the figure.
/// </summary>
public class FigureLabel : Measurement
{
    bool placed;

    IFigure Figure => Dependencies.Count > 0 ? Dependencies[0] : null;

    protected override string Kind => "Name";

    /// <summary>"of line g": "Name of line g"</summary>
    public override string Construction => Figure != null ? "of " + ConstructionText.Of(Figure) : null;

    public override void OnAddingToDrawing(Drawing drawing)
    {
        base.OnAddingToDrawing(drawing);

        // undo puts the label back without going through the figure's setter
        if (Figure is FigureBase figure && figure.NameLabel == null)
        {
            figure.NameLabel = this;
        }
    }

    public override void OnRemovingFromDrawing(Drawing drawing)
    {
        base.OnRemovingFromDrawing(drawing);
        if (Figure is FigureBase figure && figure.NameLabel == this)
        {
            figure.NameLabel = null;
        }
    }

    protected override ZOrder DefaultLayer()
    {
        return ZOrder.PointLabels;
    }

    /// <summary>Where the offset is measured from: on the figure, and for a line where the eye finds it (<see cref="FigureAnchor"/>)</summary>
    public override Point Anchor => FigureAnchor.Of(Figure, Drawing?.CoordinateSystem);

    protected override bool KeepsClear => true;

    /// <summary>A segment's name doesn't go across its segment, a circle's across its circle</summary>
    public override IEnumerable<IFigure> FiguresThroughAnchor()
    {
        return Figure != null ? new[] { Figure } : Enumerable.Empty<IFigure>();
    }

    public override void UpdateVisual()
    {
        var figure = Figure;
        if (figure == null || Drawing == null)
        {
            return;
        }

        Text = NameDisplay.Format(figure.Name);
        if (!placed && !Text.IsEmpty())
        {
            placed = true;
            Offset = FirstOffset(FigureAnchor.DefaultOffset(figure, MeasureSize(), Drawing.CoordinateSystem));
        }

        base.UpdateVisual();
        bool shown = Visible && figure.Visible;
        if (figure is LineBase && !(figure is Segment))
        {
            // a line that runs outside the window has nowhere to write its name
            shown = shown && FigureAnchor.IsOnScreen(ToPhysical(Anchor), Drawing.CoordinateSystem);
            if (shown)
            {
                KeepOnScreen();
            }
        }

        // a hidden figure keeps its name to itself
        Shape.IsVisible = shown;
    }

    /// <summary>The name pushed back in, whole, where a line's anchor by the edge of the window and the offset put it past the edge (<see cref="FigureAnchor.KeptOnScreen"/>)</summary>
    void KeepOnScreen()
    {
        var topLeft = ToPhysical(Coordinates);
        var kept = FigureAnchor.KeptOnScreen(topLeft, MeasureSize(), Drawing.CoordinateSystem);
        if (kept != topLeft && kept.Exists())
        {
            Coordinates = ToLogical(kept);
            Shape.MoveTo(kept);
        }
    }

    /// <summary>
    /// For the name of a polygon's side, which is written outside the polygon: a name that
    /// is inside goes across the side, as far from it as it was.
    /// </summary>
    /// <param name="polygon">The vertices, logical</param>
    public void KeepOutside(IList<Point> polygon)
    {
        // the default place first
        UpdateVisual();
        var size = MeasureSize();
        var half = new Point(size.Width / 2, size.Height / 2);
        var direction = RightAngleMark.Direction(new Point(), Offset + half);
        if (direction == null || Drawing == null)
        {
            return;
        }

        // a pixel from the side towards the name: the name itself may be past a narrow polygon
        var near = ToLogical(ToPhysical(Anchor) + direction.Value);
        if (polygon.IsPointInPolygon(near))
        {
            Offset = -Offset - half * 2;
            UpdateVisual();
        }
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);
        placed = true;
    }

    /// <summary>Not in the grid: the name shows and hides with its figure's Show name (see <see cref="PointLabel.Visible"/>)</summary>
    [PropertyGridVisible(false)]
    public override bool Visible
    {
        get => base.Visible; set => base.Visible = value;
    }

    [PropertyGridVisible(false)]
    public override string Text => base.Text;
}
