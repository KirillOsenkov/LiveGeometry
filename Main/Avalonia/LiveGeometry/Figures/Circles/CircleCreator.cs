using System.Collections.Generic;
using System.ComponentModel;

namespace LiveGeometry;

[Category(BehaviorCategories.Circles)]
[Order(1)]
public class CircleCreator : FigureCreator
{
    protected override IEnumerable<IFigure> CreateFigures()
    {
        yield return Factory.CreateCircle(Drawing, FoundDependencies);
    }

    protected override DependencyList InitExpectedDependencies()
    {
        return DependencyList.PointPoint;
    }

    public override string Name => "Circle";

    public override string HintText => "Click the circle center and then click a point on a circle.";

    public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
    {
        return ClickedDependencies == 1 ? "Click a point the circle goes through." : base.ConstructionHintText(args);
    }

    public override FrameworkElement CreateIcon()
    {
        return IconBuilder.BuildIcon()
            .Circle(0.5, 0.5, 0.5)
            .Point(0.5, 0.5)
            .Point(0.85, 0.15)
            .Canvas;
    }
}