using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Lines)]
    [Order(5)]
    public class PerpendicularLineCreator : FigureCreator
    {
        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreatePerpendicularLine(Drawing, FoundDependencies);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.LinePoint;
        }

        public override string Name
        {
            get
            {
                return "Perpendicular";
            }
        }

        public override string HintText
        {
            get
            {
                return "Click a line and then click a point.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Line(0, 0, 1, 1)
                .Line(0.3, 1, 1, 0.3)
                .Point(0.35, 0.35)
                .Canvas;
        }
    }
}