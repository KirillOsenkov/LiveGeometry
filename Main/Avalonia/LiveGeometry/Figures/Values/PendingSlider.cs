using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// A slider on its way into the drawing, for the tool that places it: the first click puts
/// its anchor down and from then on the knob follows the cursor, until a second click (or the
/// release of a press and drag) says where the knob starts. In between the slider is in the
/// drawing without being recorded; <see cref="Finish"/> takes it out again and hands it to
/// the tool, which adds it for real - as its undo step (the Slider tool) or as a part of one
/// (Circle by Radius, where a first click on empty paper makes the radius a slider).
/// </summary>
public class PendingSlider
{
    Drawing drawing;

    /// <summary>The slider being placed; null when none is</summary>
    public Slider Slider { get; private set; }

    public bool Exists
    {
        get { return Slider != null; }
    }

    /// <summary>The first click: the anchor goes here</summary>
    public void Start(Drawing drawing, Point coordinates)
    {
        this.drawing = drawing;
        Slider = new Slider() { Drawing = drawing, Position = coordinates };
        drawing.ActionManager.ExecuteImmediatelyWithoutRecording = true;
        Actions.Add(drawing, Slider);
        drawing.ActionManager.ExecuteImmediatelyWithoutRecording = false;
    }

    /// <summary>The knob goes where the cursor is</summary>
    public void Follow(Point coordinates)
    {
        if (Slider != null)
        {
            Slider.Value = ValueAt(coordinates);
        }
    }

    /// <summary>Whether a release here ends a press and drag, and so is the second click</summary>
    public bool IsDragged(Point coordinates)
    {
        return Slider != null
            && coordinates.Distance(Slider.Position) > 3 * drawing.CoordinateSystem.CursorTolerance;
    }

    /// <summary>
    /// The second click: the slider with its knob there, out of the drawing again, for the
    /// tool to add
    /// </summary>
    public Slider Finish(Point coordinates)
    {
        var slider = Slider;
        var value = ValueAt(coordinates);
        Cancel();
        slider.Value = value;
        return slider;
    }

    /// <summary>The slider being placed goes; nothing to undo</summary>
    public void Cancel()
    {
        if (Slider != null)
        {
            drawing.ActionManager.ExecuteImmediatelyWithoutRecording = true;
            Actions.Remove(Slider);
            drawing.ActionManager.ExecuteImmediatelyWithoutRecording = false;
            Slider = null;
        }
    }

    /// <summary>How far to the right of the anchor the cursor is; the knob can't go left of it</summary>
    double ValueAt(Point coordinates)
    {
        return System.Math.Max(0, coordinates.X - Slider.Position.X);
    }
}
