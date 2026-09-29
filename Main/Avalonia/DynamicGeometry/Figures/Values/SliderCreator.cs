using System.ComponentModel;
using Avalonia;
using Avalonia.Input;

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
    readonly PendingSlider pending = new PendingSlider();

    // Escape and right-click restart the tool (MainView, Behavior.MouseRightClick): the slider
    // being placed goes, and the construction is over for the undo button
    public override void Stopping()
    {
        if (pending.Exists)
        {
            pending.Cancel();
            RaiseConstructionComplete();
        }
    }

    public override bool IsInInitialState
    {
        get { return !pending.Exists; }
    }

    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        var coordinates = Coordinates(e);
        if (!pending.Exists)
        {
            pending.Start(Drawing, coordinates);
            Drawing.RaiseConstructionStepStarted();
            Drawing.RaiseStatusNotification("Click where the slider ends.");
            return;
        }

        Finish(coordinates);
    }

    public override void MouseMove(object sender, MouseEventArgs e)
    {
        pending.Follow(Coordinates(e));
    }

    public override void MouseUp(object sender, MouseButtonEventArgs e)
    {
        var coordinates = Coordinates(e);
        if (pending.IsDragged(coordinates))
        {
            Finish(coordinates);
        }
    }

    void Finish(Point coordinates)
    {
        var slider = pending.Finish(coordinates);
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

    /// <summary>A cross where the slider would appear, an arrow while the knob follows the cursor</summary>
    protected override Cursor GetCursor(Point coordinates)
    {
        return pending.Exists ? ArrowCursor : CrossCursor;
    }

    public override string Name
    {
        get { return "Slider"; }
    }

    public override string HintText
    {
        get
        {
            return "Click where the slider starts, then where it ends. Drag the knob at its end to change the number; click the slider where a tool asks for a length or an angle.";
        }
    }

    public override FrameworkElement CreateIcon()
    {
        return IconBuilder.BuildIcon()
            .Line(
                strokeThickness: 4,
                nameof(AppTheme.SliderTrack),
                0.1,
                0.65,
                0.9,
                0.65)
            .Point(0.1, 0.65)
            .Point(0.6, 0.65, nameof(AppTheme.PointOnFigureFill))
            .Canvas;
    }
}
