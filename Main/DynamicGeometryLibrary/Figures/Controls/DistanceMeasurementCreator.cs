using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Measure)]
    [Order(1)]
    public class DistanceMeasurementCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            var result = Factory.CreateDistanceMeasurement(Drawing, FoundDependencies);
            yield return result;
        }

        public override void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var underMouse = Drawing.Figures.HitTest(Coordinates(e),f => f is ILengthProvider && !f.DependsOn(TempPoint));
            if (underMouse != null
                && Drawing.Figures.HitTest<IPoint>(Coordinates(e)) == null)
            {
                FoundDependencies.Clear();
                FoundDependencies.Add(underMouse);
                RemoveIntermediateFigureIfNecessary();
                RemoveTempPointIfNecessary();
                AddFiguresAndRestart();
                return;
            }
            base.MouseDown(sender, e);
        }

        public override string Name
        {
            get { return "Distance"; }
        }

        public override string HintText
        {
            get
            {
                return "Click two points to measure distance between them, or a segment to measure its length.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            var builder = IconBuilder
                .BuildIcon()
                .Polygon(
                    new SolidColorBrush(Colors.Yellow),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.1, 0.8),
                    new Point(0.3, 1),
                    new Point(1, 0.3),
                    new Point(0.8, 0.1))
                .Line(0, 0.7, 0.7, 0);
            for (double i = 0.2; i <= 0.7; i += 0.1)
            {
                builder.Line(i, 0.9 - i, i + 0.1, 1 - i);
            }
            for (double i = 0.15; i <= 0.75; i += 0.1)
            {
                builder.Line(i, 0.9 - i, i + 0.05, 0.95 - i);
            }
            return builder.Canvas;
        }
    }
}