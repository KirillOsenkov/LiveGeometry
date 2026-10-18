using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;

namespace LiveGeometry;

[Category(BehaviorCategories.Misc)]
[Order(3)]
public class LabelCreator : Behavior
{
    /// <summary>
    /// A text where the click is: attached to the figure there (<see cref="AttachedLabel"/>),
    /// which it then moves with, or free on the paper
    /// </summary>
    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        var figure = FindFigureToAttach(Coordinates(e, false, false, false));
        LabelBase label;
        if (figure != null)
        {
            label = Factory.CreateAttachedLabel(Drawing, figure);
            label.Text = "Text";
        }
        else
        {
            var free = Factory.CreateLabel(Drawing);
            free.Text = "Text";
            free.MoveTo(Coordinates(e));
            label = free;
        }

        Actions.Add(Drawing, label);
        var drawing = Drawing;
        AbortAndSetDefaultTool();
        drawing.RaiseStatusNotification("");
        drawing.RaiseDisplayProperties(label, focusProperty: nameof(label.Text));
    }

    /// <summary>The figure a click attaches the text to: a point, a segment, a line, a ray, a vector, a circle or a polygon (the one chosen with Tab); none on the paper</summary>
    IFigure FindFigureToAttach(Point coordinates)
    {
        return Choice.Pick(FindFiguresToAttach(coordinates));
    }

    IReadOnlyList<IFigure> FindFiguresToAttach(Point coordinates)
    {
        return Drawing.Figures
            .HitTestAll(coordinates, figure => figure.Visible && figure.IsHitTestVisible)
            .Select(figure => Drawing.Figures.FindTopLevel(figure))
            .Where(TextAttachment.CanAttachTo)
            .Distinct()
            .ToList();
    }

    protected override IReadOnlyList<object> FindClickOptions(MouseEventArgs e)
    {
        return FindFiguresToAttach(Coordinates(e, false, false, false)).ToList<object>();
    }

    /// <summary>A halo on the figure a click would attach the text to</summary>
    protected override IFigure GetFigureToPick(MouseEventArgs e)
    {
        return FindFigureToAttach(Coordinates(e, false, false, false));
    }

    protected override Cursor GetCursor(Point coordinates)
    {
        return FindFigureToAttach(coordinates) != null ? HandCursor : base.GetCursor(coordinates);
    }

    public override string Name => "Text";

    public override string HintText => "Click the paper to add a text, or a figure to attach a text to it.";

    public override FrameworkElement CreateIcon()
    {
        var text = new TextBlock()
        {
            // as the icons of the Coordinates tab that are text (y=f(x)): same size, upright, in ink
            Text = "Abc",
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