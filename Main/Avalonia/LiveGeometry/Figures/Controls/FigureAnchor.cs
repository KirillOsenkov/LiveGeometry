using Avalonia;

namespace LiveGeometry;

/// <summary>
/// Where a label written next to a figure hangs (a figure's name, <see cref="FigureLabel"/>,
/// and a text attached to a figure, <see cref="AttachedLabel"/>, which a polygon takes in
/// its middle): an anchor on the figure,
/// a first offset from it, and for a line, which runs past the window, a place kept on screen.
/// </summary>
public static class FigureAnchor
{
    /// <summary>How far in from the edge of the window a line's label sits, in pixels</summary>
    public static double EdgeInset = 24;

    /// <summary>Between a segment and its label, in pixels</summary>
    public static double SegmentGap = 6;

    /// <summary>The least room between a line's label and the edge of the window, in pixels</summary>
    public static double EdgeMargin = 4;

    /// <summary>
    /// Where the offset is measured from, logical: a point itself, the middle of a segment, a
    /// point on the upper left of a circle, and for a line or a ray a point of its visible part
    /// near the edge of the window, where the eye finds it
    /// </summary>
    public static Point Of(IFigure figure, CoordinateSystem coordinateSystem)
    {
        switch (figure)
        {
            case IPoint point:
                return point.Coordinates;
            case Segment segment:
                return segment.Coordinates.Midpoint;
            case Ray ray:
                // the far end of what shows
                return Inward(ray.OnScreenCoordinates.P2, ray.OnScreenCoordinates.P1, coordinateSystem);
            case LineBase line:
                {
                    // the end of the visible part nearer the top of the window
                    var shown = line.OnScreenCoordinates;
                    return shown.P1.Y >= shown.P2.Y
                        ? Inward(shown.P1, shown.P2, coordinateSystem)
                        : Inward(shown.P2, shown.P1, coordinateSystem);
                }

            case ICircle circle:
                {
                    // upper left, as GeoGebra puts it
                    var center = circle.Center;
                    double radius = circle.Radius * System.Math.Sqrt(0.5);
                    return new Point(center.X - radius, center.Y + radius);
                }

            case PolygonBase polygon:
                return Centroid(polygon.VertexCoordinates);
            default:
                return figure != null ? figure.Center : new Point();
        }
    }

    /// <summary>
    /// The middle of a polygon's inside, weighted by area: the average of the vertices leans
    /// towards where they crowd, and lies outside an L. That average for one with no area.
    /// </summary>
    public static Point Centroid(Point[] vertices)
    {
        double area = 0;
        double x = 0;
        double y = 0;
        for (int i = 0; i < vertices.Length; i++)
        {
            var a = vertices[i];
            var b = vertices[(i + 1) % vertices.Length];
            double cross = a.X * b.Y - b.X * a.Y;
            area += cross;
            x += (a.X + b.X) * cross;
            y += (a.Y + b.Y) * cross;
        }

        if (System.Math.Abs(area) < 1e-12 || vertices.Length < 3)
        {
            return vertices.Midpoint();
        }

        return new Point(x / (3 * area), y / (3 * area));
    }

    /// <summary>The point <see cref="EdgeInset"/> pixels from the end towards the other end, logical</summary>
    static Point Inward(Point end, Point other, CoordinateSystem coordinateSystem)
    {
        if (coordinateSystem == null)
        {
            return end;
        }

        var from = coordinateSystem.ToPhysical(end);
        var direction = RightAngleMark.Direction(from, coordinateSystem.ToPhysical(other));
        if (direction == null)
        {
            return end;
        }

        return coordinateSystem.ToLogical(from + direction.Value * EdgeInset);
    }

    /// <summary>
    /// Where a new label goes, in pixels from the anchor: beside a segment or a line on its
    /// left (looking from its first point to its second), just inside a circle, to the right
    /// of a point (whose name goes above it), in the middle of a polygon
    /// </summary>
    public static Point DefaultOffset(IFigure figure, Size size, CoordinateSystem coordinateSystem)
    {
        if (figure is PolygonBase)
        {
            return new Point(-size.Width / 2, -size.Height / 2);
        }

        if (figure is LineBase line)
        {
            // beside the line, clear of it
            var coordinates = line is Segment ? line.Coordinates : line.OnScreenCoordinates;
            var direction = RightAngleMark.Direction(coordinateSystem.ToPhysical(coordinates.P1), coordinateSystem.ToPhysical(coordinates.P2)) ?? new Point(1, 0);
            var normal = new Point(direction.Y, -direction.X);
            double distance = SegmentGap + System.Math.Abs(normal.X) * size.Width / 2 + System.Math.Abs(normal.Y) * size.Height / 2;
            return normal * distance - new Point(size.Width / 2, size.Height / 2);
        }

        if (figure is ICircle)
        {
            return new Point(4, 2);
        }

        if (figure is PointBase point)
        {
            return new Point(point.Shape.Bounds.Width / 2 + PointLabel.Margin, -size.Height / 2);
        }

        return new Point(6, 2);
    }

    /// <summary>Whether a line's anchor is near enough to the window to write its label: a line that runs outside it has nowhere</summary>
    public static bool IsOnScreen(Point pixel, CoordinateSystem coordinateSystem)
    {
        var canvas = coordinateSystem.PhysicalSize;
        return pixel.Exists()
            && pixel.X >= -EdgeInset && pixel.X <= canvas.X + EdgeInset
            && pixel.Y >= -EdgeInset && pixel.Y <= canvas.Y + EdgeInset;
    }

    /// <summary>
    /// A line's anchor sits by the edge of the window, and the offset may put the label beyond
    /// it: the top-left corner, in pixels, that brings the label back in, whole. The offset
    /// itself is left alone, so the label finds its place again when the line comes away
    /// from the edge.
    /// </summary>
    public static Point KeptOnScreen(Point topLeft, Size size, CoordinateSystem coordinateSystem)
    {
        var canvas = coordinateSystem.PhysicalSize;
        return new Point(
            Clamp(topLeft.X, EdgeMargin, canvas.X - size.Width - EdgeMargin),
            Clamp(topLeft.Y, EdgeMargin, canvas.Y - size.Height - EdgeMargin));
    }

    static double Clamp(double value, double low, double high)
    {
        // a label wider than the window keeps its left edge in
        return high < low ? low : System.Math.Min(System.Math.Max(value, low), high);
    }
}
