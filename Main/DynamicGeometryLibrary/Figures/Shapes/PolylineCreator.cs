using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Lines)]
    [Order(12)]
    public class PolylineCreator : PolygonCreator
    {
        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreatePolyline(Drawing, FoundDependencies);
        }

        public override string Name
        {
            get { return "Polyline"; }
        }

        public override string HintText
        {
            get
            {
                return "Click points to construct a polyline. Double click or click on an existing polyline point to finish.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder
                .BuildIcon()
                .Line(0.1, 0.3, 0.4, 0.7)
                .Line(0.4, 0.7, 0.7, 0.3)
                .Line(0.7, 0.3, 1.0, 0.7)
                .Point(1.0, 0.7)
                .Point(0.7, 0.3)
                .Point(0.1, 0.3)
                .Point(0.4, 0.7)
                .Canvas;
        }
    }
}