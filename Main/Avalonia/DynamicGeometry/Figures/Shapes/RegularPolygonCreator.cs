using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Shapes)]
    [Order(4)]
    public class RegularPolygonCreator : FigureCreator
    {
        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateRegularPolygon(Drawing, FoundDependencies);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        public override string Name
        {
            get
            {
                return "Regular polygon";
            }
        }

        public override string HintText
        {
            get
            {
                return "Click the polygon center and then click a vertex.";
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
                // yellow is what the tool's two clicks make and what can be dragged: the
                // center and one vertex; the other vertices follow
                .DependentPoint(1.01, 0.63)
                .DependentPoint(0.875, 0.166)
                .DependentPoint(0.401, 0.055)
                .DependentPoint(0.068, 0.409)
                .DependentPoint(0.208, 0.874)
                .Point(0.68, 0.98)
                .Point(0.54, 0.52)
                .Canvas;
        }
    }
}