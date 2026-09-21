using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

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
            // A hatched region: all drawn, no text (text in an icon doesn't scale with it, and the
            // tab header shows a copy of this icon at another font size).
            var pentagon = new[]
            {
                new Point(0.5, 0.08),
                new Point(0.93, 0.39),
                new Point(0.77, 0.9),
                new Point(0.23, 0.9),
                new Point(0.07, 0.39)
            };

            var builder = IconBuilder
                .BuildIcon()
                .Polygon(
                    new SolidColorBrush(Color.FromArgb(255, 255, 214, 214)),
                    new SolidColorBrush(Colors.Black),
                    pentagon);

            // hatching: the parts of the lines x + y = c that are inside the pentagon
            var hatchColor = Color.FromArgb(255, 200, 96, 110);
            for (double c = 0.55; c < 1.7; c += 0.2)
            {
                var ends = new List<Point>();
                for (int i = 0; i < pentagon.Length; i++)
                {
                    var a = pentagon[i];
                    var b = pentagon[(i + 1) % pentagon.Length];
                    var sideA = a.X + a.Y - c;
                    var sideB = b.X + b.Y - c;
                    if (sideA * sideB < 0)
                    {
                        var t = sideA / (sideA - sideB);
                        ends.Add(new Point(a.X + (b.X - a.X) * t, a.Y + (b.Y - a.Y) * t));
                    }
                }

                if (ends.Count == 2)
                {
                    builder.Line(
                        hatchColor,
                        ends[0].X,
                        ends[0].Y,
                        ends[1].X,
                        ends[1].Y);
                }
            }

            return builder.Canvas;
        }
    }
}