using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;

namespace DynamicGeometry
{
    // Off the ribbon: the intersection doesn't work. The figure stays so old files load.
    [Ignore]
    [Category(BehaviorCategories.Shapes)]
    [Order(4)]
    public class PolygonIntersectionCreator : FigureCreator
    {
        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreatePolygonIntersection(Drawing, FoundDependencies);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PolygonPolygon;
        }

        public override string Name => "Polygon intersection";

        public override string HintText
        {
            get
            {
                return "Click two polygons in order to intersect them.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Polygon(
                    nameof(AppTheme.ShapeIconFill),
                    nameof(AppTheme.ShapeOutline),
                    new Point(0.68, 0.98),
                    new Point(1.01, 0.63),
                    new Point(0.875, 0.166),
                    new Point(0.401, 0.055),
                    new Point(0.068, 0.409),
                    new Point(0.208, 0.874))
                .Point(0.68, 0.98)
                .Point(1.01, 0.63)
                .Point(0.875, 0.166)
                .Point(0.401, 0.055)
                .DependentPoint(0.068, 0.409)
                .Point(0.208, 0.874)
                .DependentPoint(0.55, 0.5)
                .Canvas;
        }
    }
}