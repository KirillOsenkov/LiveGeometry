using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Lines)]
    [Order(2)]
    public class RayCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateRay(Drawing, FoundDependencies);
        }

        public override string Name
        {
            get { return "Ray"; }
        }

        public override string HintText
        {
            get
            {
                return "Click (and release) twice to create a ray.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Line(0.25, 0.75, 1, 0)
                .Point(0.25, 0.75)
                .Point(0.75, 0.25)
                .Canvas;
        }
    }
}