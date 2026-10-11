using System.Collections.Generic;
using System.Xml.Linq;
using Avalonia;

namespace DynamicGeometry;

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
    /// <summary>How far in from the edge of the window a line's name sits, in pixels</summary>
    public static double EdgeInset = 24;

    /// <summary>Between a segment and its name, in pixels</summary>
    public static double SegmentGap = 6;

    bool placed;

    IFigure Figure
    {
        get { return Dependencies.Count > 0 ? Dependencies[0] : null; }
    }

    protected override string Kind
    {
        get
        {
            return "Name";
        }
    }

    /// <summary>"of line g": "Name of line g"</summary>
    public override string Construction
    {
        get
        {
            return Figure != null ? "of " + ConstructionText.Of(Figure) : null;
        }
    }

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

    /// <summary>Where the offset is measured from: on the figure, and for a line where the eye finds it</summary>
    public override Point Anchor
    {
        get
        {
            var figure = Figure;
            switch (figure)
            {
                case Segment segment:
                    return segment.Coordinates.Midpoint;
                case Ray ray:
                    // the far end of what shows
                    return Inward(ray.OnScreenCoordinates.P2, ray.OnScreenCoordinates.P1);
                case LineBase line:
                    {
                        // the end of the visible part nearer the top of the window
                        var shown = line.OnScreenCoordinates;
                        return shown.P1.Y >= shown.P2.Y ? Inward(shown.P1, shown.P2) : Inward(shown.P2, shown.P1);
                    }

                case ICircle circle:
                    {
                        // upper left, as GeoGebra puts it
                        var center = circle.Center;
                        double radius = circle.Radius * System.Math.Sqrt(0.5);
                        return new Point(center.X - radius, center.Y + radius);
                    }

                default:
                    return figure != null ? figure.Center : new Point();
            }
        }
    }

    /// <summary>The point <see cref="EdgeInset"/> pixels from the end towards the other end, logical</summary>
    Point Inward(Point end, Point other)
    {
        if (Drawing == null)
        {
            return end;
        }

        var from = ToPhysical(end);
        var direction = RightAngleMark.Direction(from, ToPhysical(other));
        if (direction == null)
        {
            return end;
        }

        return ToLogical(from + direction.Value * EdgeInset);
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
            Offset = DefaultOffset(figure);
        }

        base.UpdateVisual();
        bool shown = Visible && figure.Visible;
        if (figure is LineBase && !(figure is Segment))
        {
            // a line that runs outside the window has nowhere to write its name
            shown = shown && IsOnScreen(ToPhysical(Anchor));
            if (shown)
            {
                KeepOnScreen();
            }
        }

        // a hidden figure keeps its name to itself
        Shape.IsVisible = shown;
    }

    bool IsOnScreen(Point pixel)
    {
        var canvas = Drawing.CoordinateSystem.PhysicalSize;
        return pixel.Exists()
            && pixel.X >= -EdgeInset && pixel.X <= canvas.X + EdgeInset
            && pixel.Y >= -EdgeInset && pixel.Y <= canvas.Y + EdgeInset;
    }

    /// <summary>The least room between a line's name and the edge of the window, in pixels</summary>
    public static double EdgeMargin = 4;

    /// <summary>
    /// A line's anchor sits by the edge of the window, and the offset may put the name beyond
    /// it: the name is pushed back in, whole. The offset itself is left alone, so the name
    /// finds its place again when the line comes away from the edge.
    /// </summary>
    void KeepOnScreen()
    {
        var canvas = Drawing.CoordinateSystem.PhysicalSize;
        var size = MeasureSize();
        var topLeft = ToPhysical(Coordinates);
        var kept = new Point(
            Clamp(topLeft.X, EdgeMargin, canvas.X - size.Width - EdgeMargin),
            Clamp(topLeft.Y, EdgeMargin, canvas.Y - size.Height - EdgeMargin));
        if (kept != topLeft && kept.Exists())
        {
            Coordinates = ToLogical(kept);
            Shape.MoveTo(kept);
        }
    }

    static double Clamp(double value, double low, double high)
    {
        // a name wider than the window keeps its left edge in
        return high < low ? low : System.Math.Min(System.Math.Max(value, low), high);
    }

    /// <summary>
    /// Where a new name goes, in pixels from the anchor: beside a segment on its left (looking
    /// from its first point to its second), just inside a circle, to the right of a line.
    /// </summary>
    Point DefaultOffset(IFigure figure)
    {
        var size = MeasureSize();
        if (figure is LineBase line)
        {
            // beside the line, clear of it, on the left looking from the first point to the second
            var coordinates = line is Segment ? line.Coordinates : line.OnScreenCoordinates;
            var direction = RightAngleMark.Direction(ToPhysical(coordinates.P1), ToPhysical(coordinates.P2)) ?? new Point(1, 0);
            var normal = new Point(direction.Y, -direction.X);
            double distance = SegmentGap + System.Math.Abs(normal.X) * size.Width / 2 + System.Math.Abs(normal.Y) * size.Height / 2;
            return normal * distance - new Point(size.Width / 2, size.Height / 2);
        }

        if (figure is ICircle)
        {
            return new Point(4, 2);
        }

        return new Point(6, 2);
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
        get
        {
            return base.Visible;
        }
        set
        {
            base.Visible = value;
        }
    }

    [PropertyGridVisible(false)]
    public override string Text
    {
        get
        {
            return base.Text;
        }
    }
}
