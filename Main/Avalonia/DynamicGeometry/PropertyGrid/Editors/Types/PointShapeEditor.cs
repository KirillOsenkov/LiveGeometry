using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("PointShape")]</c>: the shapes of a point as a
/// row of swatches, each drawn in the colors of the style being edited
/// </summary>
public class PointShapeEditorFactory : BaseValueEditorFactory<PointShapeEditor, PointShape>
{
    public PointShapeEditorFactory()
    {
        // after the plain enum editor: only the attribute chooses this one
        LoadOrder = 1;
    }
}

public class PointShapeEditor : SelectorValueEditor
{
    protected override Selector CreateSelector()
    {
        return new ListBox();
    }

    protected override void InitCore()
    {
        var style = Value.Parent as PointStyle;
        Items = Enum.GetValues<PointShape>().Select(shape => CreateSwatch(shape, style)).ToList();
        base.InitCore();
    }

    static Control CreateSwatch(PointShape shape, PointStyle style)
    {
        var marker = new PointMarker()
        {
            Kind = shape,
            Width = 14,
            Height = 14,
            Fill = (IBrush)style?.Fill ?? Brushes.White,
            Stroke = new SolidColorBrush(style?.Color ?? Colors.Black),
            StrokeThickness = 1,
            UseLayoutRounding = false,
            Margin = new Thickness(8)
        };
        var result = new Grid() { Tag = shape };
        result.Children.Add(marker);
        ToolTip.SetTip(result, shape.ToString());
        return result;
    }

    protected override ValidationResult Validate(object value)
    {
        var result = new ValidationResult();
        if ((value as Control)?.Tag is PointShape shape)
        {
            result.IsValid = true;
            result.Value = shape;
        }

        return result;
    }

    public override void UpdateEditor()
    {
        var value = GetValue();
        ShowSelected(item => value != null && Equals(((Control)item).Tag, value));
    }
}
