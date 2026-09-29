using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using AvaloniaShapes = Avalonia.Controls.Shapes;

namespace DynamicGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("SegmentDecoration")]</c>: the marks a segment can
/// wear as a row of swatches, each a short segment with the mark on it
/// </summary>
public class SegmentDecorationEditorFactory : BaseValueEditorFactory<SegmentDecorationEditor, SegmentDecoration>
{
    public SegmentDecorationEditorFactory()
    {
        // after the plain enum editor: only the attribute chooses this one
        LoadOrder = 1;
    }
}

public class SegmentDecorationEditor : SelectorValueEditor
{
    const double SwatchWidth = 34;
    const double SwatchHeight = 18;

    protected override Selector CreateSelector()
    {
        return new ListBox();
    }

    protected override void InitCore()
    {
        Items = Enum.GetValues<SegmentDecoration>().Select(CreateSwatch).ToList();
        base.InitCore();
    }

    static Control CreateSwatch(SegmentDecoration decoration)
    {
        var middle = new Point(SwatchWidth / 2, SwatchHeight / 2);
        var geometry = SegmentDecorationMark.CreateGeometry(decoration, middle, new Point(1, 0), thickness: 1.5);
        geometry.Figures.Insert(0, new PathFigure()
        {
            StartPoint = new Point(0, middle.Y),
            IsClosed = false,
            IsFilled = false,
            Segments = new PathSegments() { new LineSegment() { Point = new Point(SwatchWidth, middle.Y) } }
        });
        var path = new AvaloniaShapes.Path()
        {
            Data = geometry,
            Width = SwatchWidth,
            Height = SwatchHeight,
            StrokeThickness = 1.5,
            StrokeLineCap = PenLineCap.Round,
            StrokeJoin = PenLineJoin.Round,
            Margin = new Thickness(4, 8)
        };
        path.BindTheme(AvaloniaShapes.Shape.StrokeProperty, nameof(AppTheme.Ink));
        var result = new Grid() { Tag = decoration };
        result.Children.Add(path);
        ToolTip.SetTip(result, Caption(decoration));
        return result;
    }

    static string Caption(SegmentDecoration decoration)
    {
        switch (decoration)
        {
            case SegmentDecoration.None:
                return "No mark";
            case SegmentDecoration.OneTick:
                return "One tick";
            case SegmentDecoration.TwoTicks:
                return "Two ticks";
            case SegmentDecoration.ThreeTicks:
                return "Three ticks";
            case SegmentDecoration.OneArrow:
                return "One arrow";
            case SegmentDecoration.TwoArrows:
                return "Two arrows";
            case SegmentDecoration.ThreeArrows:
                return "Three arrows";
            default:
                return "Wave";
        }
    }

    protected override ValidationResult Validate(object value)
    {
        var result = new ValidationResult();
        if ((value as Control)?.Tag is SegmentDecoration decoration)
        {
            result.IsValid = true;
            result.Value = decoration;
        }

        return result;
    }

    public override void UpdateEditor()
    {
        var value = GetValue();
        foreach (Control item in Items)
        {
            if (Equals(item.Tag, value))
            {
                guard = true;
                Selector.SelectedItem = item;
                guard = false;
            }
        }
    }
}
