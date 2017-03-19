using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Misc)]
    [Order(1)]
    public class BezierCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPointPointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateBezier(Drawing, FoundDependencies);
        }

        public override string Name
        {
            get { return "Bezier"; }
        }

        public override string HintText
        {
            get
            {
                return "Click four points to draw a cubic bezier curve.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Bezier(0, 0.75, 0, 0, 0.5, 1, 1, 0)
                //.Line(0, 0, 0, 0.75)
                //.Line(0.5, 1, 1, 0)
                //.Point(0, 0)
                //.Point(0, 0.75)
                //.Point(1, 0)
                //.Point(0.5, 1)
                .Canvas;
        }
    }
}