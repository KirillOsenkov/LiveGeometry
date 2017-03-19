using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Lines)]
    [Order(7)]
    public class AngleBisectorCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPointPoint;
        }

        public override void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var underMouse = Drawing.Figures.HitTest(Coordinates(e));
            if (underMouse != null
                && (underMouse is AngleArc || underMouse is AngleMeasurement))
            {
                FoundDependencies.AddRange(underMouse.Dependencies);
            }
            base.MouseDown(sender, e);
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            var result = Factory.CreateAngleBisector(Drawing, FoundDependencies);
            yield return result;
        }

        public override string Name
        {
            get { return "Angle Bisector"; }
        }

        public override string HintText
        {
            get
            {
                return "Click an angle vertex, then click two points on the angle sides to create an angle bisector. You can also click an angle measurement.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            const double a = 0.9, b = 0.1;
            var builder = IconBuilder.BuildIcon()
                .Line(a, a, b, a)
                .Line(b, a, b, b)
                .Line(a, b, b, a)
                .Arc(b, a, 0.4, a, b, 0.6)
                .Point(a, a)
                .Point(b, a)
                .Point(b, b);

            return builder.Canvas;
        }
    }
}