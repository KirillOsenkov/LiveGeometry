namespace DynamicGeometry;

/// <summary>
/// Where a <see cref="Label"/> is held: nowhere (a place in the plane, like any figure), or a
/// corner of the canvas, from which it keeps its distance in pixels while the plane zooms and
/// pans underneath - a caption, a heading, a readout.
/// </summary>
public enum LabelPin
{
    None,
    TopLeft,
    TopRight,
    BottomLeft,
    BottomRight
}
