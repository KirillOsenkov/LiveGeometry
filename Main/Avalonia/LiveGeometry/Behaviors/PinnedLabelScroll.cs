using System.Linq;
using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// The pinned labels of a drawing as one thing to drag along with the view
/// (<see cref="Drawing.FixedLabels"/>): moved by a logical offset, they move as far on the
/// screen as the view does when it is dragged by that offset. By pixels, so that undoing the
/// drag, which moves it back by the same offset, puts the text back where it was.
/// </summary>
public class PinnedLabelScroll : IMovable
{
    readonly Drawing drawing;

    public PinnedLabelScroll(Drawing drawing)
    {
        this.drawing = drawing;
    }

    /// <summary>How far it was moved in all, in logical units</summary>
    public Point Coordinates { get; private set; }

    public bool AllowMove()
    {
        return true;
    }

    public void MoveTo(Point position)
    {
        var offset = position - Coordinates;
        Coordinates = position;

        // the plane's y goes up, the screen's down
        double unitLength = drawing.CoordinateSystem.UnitLength;
        var pixels = new Avalonia.Vector(offset.X * unitLength, -offset.Y * unitLength);
        foreach (var label in drawing.Figures.OfType<Label>().Where(label => label.Pin != LabelPin.None))
        {
            label.ScrollPinned(pixels);
        }
    }
}
