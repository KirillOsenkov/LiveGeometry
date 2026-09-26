using System.ComponentModel;
using Avalonia;
using Avalonia.Input;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Two clicks: where the slider sits (its anchor), then where its knob starts - the knob
/// follows the cursor in between, so the second click also shows which way it slides. A
/// press, drag and release does the same in one motion, like a segment. The slider is a
/// single figure, so adding it is the undo step; no transaction.
/// </summary>
[Category(BehaviorCategories.Measure)]
[Order(4)]
public class SliderCreator : Behavior
{
    // in the drawing between the clicks, not recorded, so that its knob can follow the cursor
    Slider pending;
    Point anchorClick;

    // Escape and right-click restart the tool (MainView, Behavior.MouseRightClick): the slider
    // being placed goes, and the construction is over for the undo button
    public override void Stopping()
    {
        if (pending != null)
        {
            RemovePending();
            RaiseConstructionComplete();
        }
    }

    public override bool IsInInitialState
    {
        get { return pending == null; }
    }

    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        var coordinates = Coordinates(e);
        if (pending == null)
        {
            anchorClick = coordinates;
            pending = new Slider() { Drawing = Drawing, Position = coordinates };
            Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = true;
            Actions.Add(Drawing, pending);
            Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = false;
            Drawing.RaiseConstructionStepStarted();
            Drawing.RaiseStatusNotification("Click where the knob starts.");
            return;
        }

        Finish(coordinates);
    }

    public override void MouseMove(object sender, MouseEventArgs e)
    {
        if (pending != null)
        {
            pending.Value = ValueAt(Coordinates(e));
        }
    }

    public override void MouseUp(object sender, MouseButtonEventArgs e)
    {
        var coordinates = Coordinates(e);
        if (pending != null && coordinates.Distance(anchorClick) > 3 * CursorTolerance)
        {
            Finish(coordinates);
        }
    }

    /// <summary>How far to the right of the anchor the cursor is; the knob can't go left of it</summary>
    double ValueAt(Point coordinates)
    {
        return System.Math.Max(0, coordinates.X - pending.Position.X);
    }

    void Finish(Point coordinates)
    {
        var slider = pending;
        var value = ValueAt(coordinates);
        RemovePending();
        slider.Value = value;
        Actions.Add(Drawing, slider);
        RaiseConstructionComplete();
        Drawing.RaiseDisplayProperties(slider);
        Drawing.RaiseStatusNotification(slider.Name + ": drag the knob, or type its value in the panel.");
    }

    void RaiseConstructionComplete()
    {
        Drawing.RaiseConstructionStepComplete(new Drawing.ConstructionStepCompleteEventArgs()
        {
            ConstructionComplete = true
        });
    }

    void RemovePending()
    {
        if (pending != null)
        {
            Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = true;
            Actions.Remove(pending);
            Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = false;
            pending = null;
        }
    }

    /// <summary>A cross where the slider would appear, an arrow while the knob follows the cursor</summary>
    protected override Cursor GetCursor(Point coordinates)
    {
        return pending == null ? CrossCursor : ArrowCursor;
    }

    public override string Name
    {
        get { return "Slider"; }
    }

    public override string HintText
    {
        get
        {
            return "Click where the slider goes, then where its knob starts. Drag the knob to change the number; click the slider where a tool asks for a length or an angle.";
        }
    }

    public override FrameworkElement CreateIcon()
    {
        return IconBuilder.BuildIcon()
            .Line(0.1, 0.65, 0.9, 0.65)
            .Point(0.1, 0.65)
            .Point(0.6, 0.65, new SolidColorBrush(Color.FromArgb(255, 124, 227, 139)))
            .Canvas;
    }
}
