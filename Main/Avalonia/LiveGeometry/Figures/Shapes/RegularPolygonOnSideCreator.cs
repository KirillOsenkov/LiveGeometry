using System.ComponentModel;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// A regular polygon on a side: two adjacent vertices, the polygon to the left of the first
/// to the second, with as many sides as the tool's panel says - the Square with any number
/// of sides, for tilings. Its panels are the Regular polygon tool's, the count shared.
/// </summary>
[Category(BehaviorCategories.Shapes)]
[Order(5)]
public class RegularPolygonOnSideCreator : RegularPolygonCreator
{
    protected override bool OnSide => true;

    public override string Name
    {
        get
        {
            return "Polygon on a side";
        }
    }

    public override string HintText
    {
        get
        {
            return "Click two adjacent vertices: the regular polygon goes to the left of the first to the second.";
        }
    }

    public override FrameworkElement CreateIcon()
    {
        // a hexagon on its bottom side; yellow is what the tool's two clicks make
        double a = 0.25;
        double b = 0.75;
        double rise = 0.433;
        return IconBuilder.BuildIcon()
            .Polygon(
                nameof(AppTheme.ShapeIconFill),
                nameof(AppTheme.ShapeOutline),
                new Point(a, 0.95),
                new Point(b, 0.95),
                new Point(1, 0.95 - rise),
                new Point(b, 0.95 - 2 * rise),
                new Point(a, 0.95 - 2 * rise),
                new Point(0, 0.95 - rise))
            .DependentPoint(1, 0.95 - rise)
            .DependentPoint(b, 0.95 - 2 * rise)
            .DependentPoint(a, 0.95 - 2 * rise)
            .DependentPoint(0, 0.95 - rise)
            .Point(a, 0.95)
            .Point(b, 0.95)
            .Canvas;
    }
}
