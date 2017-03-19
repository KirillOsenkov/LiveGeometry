using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Lines)]
    [Order(1)]
    public class SegmentCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateSegment(Drawing, FoundDependencies);
        }

        public override string Name
        {
            get { return "Segment"; }
        }

        public override string HintText
        {
            get
            {
                return "Click (and release) twice to connect two points with a segment.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Line(0.25, 0.75, 0.75, 0.25)
                .Point(0.25, 0.75)
                .Point(0.75, 0.25)
                .Canvas;
        }
    }
}