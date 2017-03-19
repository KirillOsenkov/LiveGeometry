using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Measure)]
    [Order(3)]
    public class AreaMeasurementCreator : FigureCreator
    {
        protected override void Click(Point coordinates)
        {
            var figure = Drawing.Figures.HitTest(coordinates);
            if (figure is IShapeWithInterior)
            {
                FoundDependencies.Clear();
                FoundDependencies.Add(figure);
                RemoveIntermediateFigureIfNecessary();
                RemoveTempPointIfNecessary();
                AddFiguresAndRestart();
                return;
            }

            //var point = Drawing.Figures.HitTest<IPoint>(coordinates);
            if (figure is PointBase && FoundDependencies.Count >= 4 // 4 including the TempPoint 
                // (and 3 after TempPoint is removed)
                && FoundDependencies.Contains(figure))
            {
                RemoveIntermediateFigureIfNecessary();
                RemoveTempPointIfNecessary();
                AddFiguresAndRestart();
                return;
            }

            base.Click(coordinates);
        }

        protected override bool CanCreateTempResults()
        {
            return false;
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return null;
        }

        protected override System.Type GetExpectedDependencyType()
        {
            return typeof(IPoint);
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateAreaMeasurement(Drawing, FoundDependencies);
        }

        protected override IFigure CreateIntermediateFigure()
        {
            if (!FoundDependencies.All(f => f is IPoint))
            {
                return null;
            }
            else if (FoundDependencies.Count >= 3)
            {
                var result = Factory.CreateAreaMeasurement(Drawing, FoundDependencies);
                return result;
            }
            return null;
        }

        public override string Name
        {
            get { return "Area"; }
        }

        public override string HintText
        {
            get
            {
                return "Click a polygon, ellipse, circle, or a list of points to measure its area.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            var result = IconBuilder
                .BuildIcon()
                .Polygon(
                    new SolidColorBrush(Color.FromArgb(255, 255, 200, 200)),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.2, 0.4),
                    new Point(0.3, 0.8),
                    new Point(0.7, 0.8),
                    new Point(0.8, 0.4),
                    new Point(0.6, 0.2));
            var text = new TextBlock() { Text = "S²" };
            var canvas = result.Canvas;
            canvas.Children.Add(text);
            Canvas.SetLeft(text, canvas.Width / 2.0 - text.ActualWidth / 2.0);
            Canvas.SetTop(text, canvas.Height / 2.0 - text.ActualHeight / 2.0);
            return canvas;
        }
    }
}