using System.Collections.Generic;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// The name of a line, ray, segment or circle, written next to it (as a <see cref="PointLabel"/>
/// is a point's name): a label that depends on the figure and sits <see cref="LabelWithOffset.Offset"/>
/// pixels from an anchor on it - the middle of a segment, a point on the upper left of a
/// circle, and for a line or a ray a point of its visible part near the edge of the window,
/// so that the name is always on screen however far the line runs (<see cref="AnchoredLabel"/>).
/// Shown and hidden through the figure's "Show name" (<see cref="FigureBase.HasNameLabel"/>),
/// which adds and removes it; not in the Figure List, since it is part of the figure.
/// </summary>
public class FigureLabel : NameLabel
{
    protected override string Kind => "Name";

    public override void UpdateVisual()
    {
        if (Figure != null)
        {
            Text = NameDisplay.Format(Figure.Name);
        }

        base.UpdateVisual();
    }

    /// <summary>
    /// For the name of a polygon's side, which is written outside the polygon: a name that
    /// is inside goes across the side, as far from it as it was.
    /// </summary>
    /// <param name="polygon">The vertices, logical</param>
    public void KeepOutside(IList<Point> polygon)
    {
        // the default place first
        PlaceIfNew();
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
}
