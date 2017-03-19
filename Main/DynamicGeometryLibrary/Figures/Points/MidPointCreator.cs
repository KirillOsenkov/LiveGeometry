using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Points)]
    [Order(2)]
    public class MidpointCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            MidPoint result = Factory.CreateMidPoint(Drawing, FoundDependencies);
            yield return result;
        }

        public override void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var underMouse = Drawing.Figures.HitTest<Segment>(Coordinates(e));
            if (underMouse != null
                && underMouse.Dependencies.Count() == 2
                && Drawing.Figures.HitTest<IPoint>(Coordinates(e)) == null)
            {
                FoundDependencies.AddRange(underMouse.Dependencies);
            }
            base.MouseDown(sender, e);
        }

        public override string Name
        {
            get { return "Midpoint"; }
        }

        public override string HintText
        {
            get
            {
                return "Click two points (or a segment) to construct a midpoint.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Point(0.25, 0.75)
                .DependentPoint(0.5, 0.5)
                .Point(0.75, 0.25)
                .Canvas;
        }
    }
}