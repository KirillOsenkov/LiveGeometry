using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// Where a <see cref="Label"/> (or a <see cref="ShowHideControl"/>) is held: nowhere (a place
/// in the plane, like any figure), or a corner of the canvas, from which it keeps its distance
/// in pixels while the plane zooms and pans underneath - a caption, a heading, a readout.
/// </summary>
public enum LabelPin
{
    None,
    TopLeft,
    TopRight,
    BottomLeft,
    BottomRight
}

/// <summary>The pixels of a pin, for whatever can be pinned</summary>
public static class Pinning
{
    /// <summary>
    /// The top-left corner of something of this size whose pinned corner sits offset pixels
    /// inward from the same corner of the canvas
    /// </summary>
    public static Point TopLeft(LabelPin pin, Point offset, Size size, Point canvas)
    {
        double x = pin == LabelPin.TopLeft || pin == LabelPin.BottomLeft
            ? offset.X
            : canvas.X - offset.X - size.Width;
        double y = pin == LabelPin.TopLeft || pin == LabelPin.TopRight
            ? offset.Y
            : canvas.Y - offset.Y - size.Height;
        return new Point(x, y);
    }

    /// <summary>The offset that puts something of this size at this top-left corner</summary>
    public static Point OffsetFrom(LabelPin pin, Point topLeft, Size size, Point canvas)
    {
        double x = pin == LabelPin.TopLeft || pin == LabelPin.BottomLeft
            ? topLeft.X
            : canvas.X - topLeft.X - size.Width;
        double y = pin == LabelPin.TopLeft || pin == LabelPin.TopRight
            ? topLeft.Y
            : canvas.Y - topLeft.Y - size.Height;
        return new Point(x, y);
    }
}
