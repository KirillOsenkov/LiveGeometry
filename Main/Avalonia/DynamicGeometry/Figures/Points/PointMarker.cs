using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls.Shapes;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>What a point looks like when its style has no character (<see cref="PointStyle.Shape"/>)</summary>
public enum PointShape
{
    Circle,
    Triangle,
    Square,
    Diamond,
    Pentagon,
    Hexagon
}

/// <summary>
/// The visual of a point: a <see cref="PointShape"/> filling its Width x Height, centered on
/// the point, drawn with Fill and Stroke like any shape. Or a character (an emoji) instead,
/// which an <see cref="EmojiGlyph"/> inside draws (Shape.Render is sealed); one that isn't a
/// color emoji takes the Fill.
/// </summary>
public class PointMarker : Shape
{
    public static readonly StyledProperty<PointShape> KindProperty =
        AvaloniaProperty.Register<PointMarker, PointShape>(nameof(Kind));

    public static readonly StyledProperty<string> CharacterProperty =
        AvaloniaProperty.Register<PointMarker, string>(nameof(Character));

    public static readonly StyledProperty<bool> IsHighlightedProperty =
        AvaloniaProperty.Register<PointMarker, bool>(nameof(IsHighlighted));

    static PointMarker()
    {
        // Width and Height, not Bounds: Bounds changes with every move of the point
        AffectsGeometry<PointMarker>(WidthProperty, HeightProperty, StrokeThicknessProperty, KindProperty, CharacterProperty);
    }

    public PointShape Kind
    {
        get => GetValue(KindProperty);
        set => SetValue(KindProperty, value);
    }

    /// <summary>Drawn instead of the shape when not empty</summary>
    public string Character
    {
        get => GetValue(CharacterProperty);
        set => SetValue(CharacterProperty, value);
    }

    /// <summary>The point is selected: a character gets a plate behind it (a shape only grows)</summary>
    public bool IsHighlighted
    {
        get => GetValue(IsHighlightedProperty);
        set => SetValue(IsHighlightedProperty, value);
    }

    EmojiGlyph glyph;

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == CharacterProperty
            || change.Property == FillProperty
            || change.Property == IsHighlightedProperty)
        {
            UpdateGlyph();
        }
    }

    void UpdateGlyph()
    {
        if (string.IsNullOrEmpty(Character))
        {
            if (glyph != null)
            {
                VisualChildren.Remove(glyph);
                glyph = null;
            }

            return;
        }

        if (glyph == null)
        {
            glyph = new EmojiGlyph() { IsHitTestVisible = false };
            VisualChildren.Add(glyph);
            InvalidateMeasure();
        }

        glyph.Text = Character;
        glyph.Foreground = Fill;
        glyph.IsHighlighted = IsHighlighted;
    }

    protected override Size MeasureOverride(Size availableSize)
    {
        glyph?.Measure(availableSize);
        return base.MeasureOverride(availableSize);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        var result = base.ArrangeOverride(finalSize);
        if (glyph == null)
        {
            return result;
        }

        // without geometry the shape would take no room, and sit in the middle of its place
        glyph.Arrange(new Rect(finalSize));
        return finalSize;
    }

    /// <summary>
    /// How far the corners of each polygon reach, in radii of the circle: inscribed in the
    /// circle they would look smaller than it, so they reach out until they look about as big.
    /// </summary>
    static readonly Dictionary<PointShape, double> Reach = new Dictionary<PointShape, double>()
    {
        [PointShape.Triangle] = 1.3,
        [PointShape.Square] = 1.2,
        [PointShape.Diamond] = 1.3,
        [PointShape.Pentagon] = 1.15,
        [PointShape.Hexagon] = 1.1
    };

    protected override Geometry CreateDefiningGeometry()
    {
        // the glyph draws the character
        if (!string.IsNullOrEmpty(Character))
        {
            return null;
        }

        double width = double.IsNaN(Width) ? 0 : Width;
        double height = double.IsNaN(Height) ? 0 : Height;
        var center = new Point(width / 2, height / 2);
        double radius = System.Math.Max(0, System.Math.Min(width, height) / 2 - StrokeThickness / 2);
        return Kind switch
        {
            PointShape.Triangle => RegularPolygon(center, radius * Reach[Kind], sides: 3, startAngle: -90, squeeze: 1),
            PointShape.Square => RegularPolygon(center, radius * Reach[Kind], sides: 4, startAngle: 45, squeeze: 1),
            PointShape.Diamond => RegularPolygon(center, radius * Reach[Kind], sides: 4, startAngle: -90, squeeze: 0.75),
            PointShape.Pentagon => RegularPolygon(center, radius * Reach[Kind], sides: 5, startAngle: -90, squeeze: 1),
            PointShape.Hexagon => RegularPolygon(center, radius * Reach[Kind], sides: 6, startAngle: 0, squeeze: 1),
            _ => new EllipseGeometry(new Rect(center.X - radius, center.Y - radius, 2 * radius, 2 * radius))
        };
    }

    /// <param name="startAngle">Where the first corner is, in degrees clockwise from the right (y down)</param>
    /// <param name="squeeze">The width in heights (a diamond is narrower than it is tall)</param>
    static Geometry RegularPolygon(Point center, double radius, int sides, double startAngle, double squeeze)
    {
        var figure = new PathFigure()
        {
            IsClosed = true,
            IsFilled = true,
            Segments = new PathSegments()
        };
        for (int i = 0; i < sides; i++)
        {
            double angle = (startAngle + 360.0 * i / sides) * System.Math.PI / 180;
            var corner = new Point(
                center.X + radius * squeeze * System.Math.Cos(angle),
                center.Y + radius * System.Math.Sin(angle));
            if (i == 0)
            {
                figure.StartPoint = corner;
            }
            else
            {
                figure.Segments.Add(new LineSegment() { Point = corner });
            }
        }

        return new PathGeometry() { Figures = new PathFigures() { figure } };
    }
}
