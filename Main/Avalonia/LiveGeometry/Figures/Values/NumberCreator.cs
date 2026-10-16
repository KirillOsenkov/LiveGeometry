using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;

namespace LiveGeometry;

/// <summary>
/// One click: where the number goes. The side panel opens with its Formula box, so a value
/// (3) or a formula (d.ArcLength + f.ArcLength) is typed next; a click on a figure, or on
/// its row in the Figure List, puts it into the formula while the box has the keyboard
/// (<see cref="ReferencePicker"/>). The number is a single figure, so adding it is the undo
/// step; the formula typed after is an edit of its own.
/// </summary>
[Category(BehaviorCategories.Measure)]
[Order(4)]
public class NumberCreator : Behavior
{
    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        var drawing = Drawing;
        var number = new Number() { Drawing = drawing };
        number.Position = Coordinates(e);
        number.Visible = true;
        Actions.Add(drawing, number);
        AbortAndSetDefaultTool();
        drawing.RaiseStatusNotification("");
        drawing.RaiseDisplayProperties(number, focusProperty: Number.FormulaName);
    }

    /// <summary>A cross: a new number appears here</summary>
    protected override Cursor GetCursor(Point coordinates)
    {
        return CrossCursor;
    }

    public override string Name => "Number";

    public override string HintText => "Click where the number goes, then type a value or a formula. Click a figure, or its row in the Figure List, to put it into the formula.";

    public override FrameworkElement CreateIcon()
    {
        var text = new TextBlock()
        {
            // as the Text tool's "Abc": a caption as the number shows it, upright, in ink
            Text = "a=2",
            FontSize = 13,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        text.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Ink));
        var grid = new Grid()
        {
            MinWidth = IconBuilder.IconSize,
            MinHeight = IconBuilder.IconSize,
        };
        grid.Children.Add(text);
        return grid;
    }
}
