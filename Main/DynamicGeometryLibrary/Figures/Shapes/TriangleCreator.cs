using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Shapes)]
    [Order(1)]
    public class TriangleCreator : ShapeCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreatePolygon(Drawing, FoundDependencies);

            for (int i = 0; i < FoundDependencies.Count; i++)
            {
                // get two consecutive vertices of the polygon
                int j = (i + 1) % FoundDependencies.Count;
                IPoint p1 = FoundDependencies[i] as IPoint;
                IPoint p2 = FoundDependencies[j] as IPoint;
                // try to find if there is already a line connecting them
                if (Drawing.Figures.FindLine(p1, p2) == null)
                {
                    // if not, create a new segment
                    var segment = Factory.CreateSegment(Drawing, new[] { p1, p2 });
                    yield return segment;
                }
            }
        }

        protected override IFigure CreateIntermediateFigure()
        {
            if (FoundDependencies.Count == 2)
            {
                return Factory.CreateSegment(Drawing, FoundDependencies);
            }
            return null;
        }

        public override string Name
        {
            get { return "Triangle"; }
        }

        public override string HintText
        {
            get
            {
                return "Click 3 points to construct a triangle.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder
                .BuildIcon()
                .Polygon(
                    Factory.CreateDefaultFillBrush(),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.2, 0.9),
                    new Point(0.5, 0.1),
                    new Point(0.9, 0.9))
                .Canvas;
        }
    }
}