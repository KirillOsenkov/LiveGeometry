using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Markup;
using System.Windows.Shapes;
using M = System.Math;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Measure)]
    [Order(2)]
    public class AngleMeasurementCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            var result = Factory.CreateAngleMeasurement(Drawing, FoundDependencies);
            yield return result;
        }

        protected override void CreateTempResults()
        {
            base.CreateTempResults();
            Actions.Add(Drawing, Factory.CreateAngleArc(Drawing, FoundDependencies));
        }

        public override string Name
        {
            get { return "Angle"; }
        }

        public override string HintText
        {
            get
            {
                return "Click an angle vertex, and then click two points on the angle sides to measure the angle.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            var builder = IconBuilder.BuildIcon();
            var size = IconBuilder.IconSize;
            var centerX = (size - 4) / (2 * size);
            var centerY = 1 - 4 / size;
            builder.Line(centerX, centerY, 0.8, 0.2)
                .Line(centerX, centerY, 1, centerY);

            var xaml = string.Format(@"
  <Path xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    Stroke='Black'
    StrokeThickness='1'
    Fill='#33FF33'
    Data='m 0,{0} v-4 a {1},{1} 0 0 1 {2},0 v4 z m {4},-5 a 6,6 0 0 1 {3},0 z'
  />", size, (size - 4) / 2, size - 4, size / 2, (size - 4) / 4 - 1);
#if SILVERLIGHT
            var path = XamlReader.Load(xaml) as Path;
#else
            var path = XamlReader.Parse(xaml) as Path;
#endif
            builder.Canvas.Children.Add(path);

            var radius = ((size - 4) / 2) * 0.95;
            var radiusSmall = radius * 0.8;
            Point center = new Point(radius + 1, size - 4);
            for (double i = 0; i < 16; i++)
            {
                var angle = i * Math.PI / 15;
                builder.Line((center.X + radius * M.Cos(angle)) / size,
                    (center.Y - radius * M.Sin(angle)) / size,
                    (center.X + radiusSmall * M.Cos(angle)) / size,
                    (center.Y - radiusSmall * M.Sin(angle)) / size);
            }

            return builder.Canvas;
        }
    }
}