using System.Collections.Generic;
using System.ComponentModel;

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
            // the ends at the bottom, each pulled up and away by its handle
            return IconBuilder.BuildIcon()
                .DashedLine(nameof(AppTheme.Ink), 0.1, 0.67, 0.22, 0.2)
                .DashedLine(nameof(AppTheme.Ink), 0.53, 0.8, 0.9, 0.43)
                .Bezier(
                    strokeThickness: 2,
                    nameof(AppTheme.LineAccent),
                    0.1, 0.67, 0.22, 0.2, 0.9, 0.43, 0.53, 0.8)
                .Point(0.22, 0.2)
                .Point(0.9, 0.43)
                .Point(0.1, 0.67)
                .Point(0.53, 0.8)
                .Canvas;
        }
    }
}