using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;

namespace DynamicGeometry;

[Category(BehaviorCategories.Shapes)]
[Order(2)]
public class SquareCreator : ShapeCreator
{
    // the square of the construction just made, new or one the drawing had on its corners
    Polygon square;

    protected override IEnumerable<IFigure> CreateFigures()
    {
        var p1 = FoundDependencies[0] as IPoint;
        var p2 = FoundDependencies[1] as IPoint;
        if (p1.Coordinates.X == p2.Coordinates.X && p1.Coordinates.Y == p2.Coordinates.Y)
        {
            var p2AsMovable = p2 as IMovable;
            if (p2AsMovable != null)
            {
                p2AsMovable.MoveTo(p2.Coordinates.Plus(new Point(.01, 0)));
            }
        }

        // The other two corners a quarter turn counterclockwise from p1 -> p2, so that a
        // square drawn left to right stands on its first side: points by coordinates over
        // the two (C = A.Y + B.X - B.Y, -A.X + B.X + B.Y), where a circle, a perpendicular,
        // their crossing, a midpoint and a reflection once made them. A side drawn between
        // the two points already (segment AB, then a square on A and B) is the square's
        // first side, as with the Triangle and Polygon tools: a second segment lay on top
        // of it as AB2. A corner the drawing has is taken too, with the sides to it.
        return RegularPolygonBuilder.Create(Drawing, p1, p2, sides: 4, onSide: true, out square);
    }

    /// <summary>The base side's panel, as for any figure with a length; nothing but a word when the square was there already</summary>
    protected override void ShowCreatedFigure(IList<IFigure> figures)
    {
        if (square != null && !figures.Contains(square))
        {
            Drawing.RaiseStatusNotification(square.Title + " is there already.");
            return;
        }

        base.ShowCreatedFigure(figures);
    }

    protected override DependencyList InitExpectedDependencies()
    {
        return DependencyList.PointPoint;
    }

    public override string Name
    {
        get { return "Square"; }
    }

    public override FrameworkElement CreateIcon()
    {
        double a = 0.2, b = 0.8;
        return IconBuilder.BuildIcon()
            .Polygon(
                nameof(AppTheme.ShapeIconFill),
                nameof(AppTheme.ShapeOutline),
                new Point(a, a),
                new Point(b, a),
                new Point(b, b),
                new Point(a, b))
            .DependentPoint(a, a)
            .DependentPoint(b, a)
            .Point(b, b)
            .Point(a, b)
            .Canvas;
    }

    public override string HintText
    {
        get
        {
            return "Create a square given two adjacent vertices";
        }
    }
}