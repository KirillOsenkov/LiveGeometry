using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Lines)]
    [Order(4)]
    public class ParallelLineCreator : FigureCreator
    {
        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateParallelLine(Drawing, FoundDependencies);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.LinePoint;
        }

        public override string Name
        {
            get
            {
                return "Parallel";
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
                .Line(0, 0.7, 0.7, 0)
                .Line(0.3, 1, 1, 0.3)
                .Point(0.35, 0.35)
                .Canvas;
        }
    }
}