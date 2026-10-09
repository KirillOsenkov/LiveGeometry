using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Lines)]
    [Order(4)]
    public class VectorCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateVector(Drawing, FoundDependencies);
        }

        public override string Name
        {
            get { return "Vector"; }
        }

        public override string HintText
        {
            get
            {
                return "Click (and release) twice to connect two points with a vector.";
            }
        }

        public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
        {
            return ClickedDependencies == 1 ? "Click the end of the vector, where its arrowhead goes." : base.ConstructionHintText(args);
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Line(0.25, 0.75, 0.75, 0.25)
                .Polygon(nameof(AppTheme.Ink), nameof(AppTheme.Ink), new Point(0.75, 0.25), new Point(0.5, 0.4), new Point(0.6, 0.5))
                .Canvas;
        }
    }
}