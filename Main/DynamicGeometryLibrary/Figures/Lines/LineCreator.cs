using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Lines)]
    [Order(3)]
    public class LineTwoPointsCreator : FigureCreator
    {
        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateLineTwoPoints(Drawing, FoundDependencies);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        public override string Name
        {
            get
            {
                return "Line";
            }
        }

        public override string HintText
        {
            get
            {
                return "Click (and release) twice to draw a line between two points.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Line(0, 1, 1, 0)
                .Point(0.25, 0.75)
                .Point(0.75, 0.25)
                .Canvas;
        }
    }
}