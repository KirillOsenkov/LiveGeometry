using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Shapes)]
    [Order(3)]
    public class PolygonCreator : ShapeCreator
    {
        protected int FoundDependenciesMinimum = 3; // Includes TempPoint
        protected override void Click(Point coordinates)
        {
            var point = Drawing.Figures.HitTest<IPoint>(coordinates);
            if (point != null
                && FoundDependencies.Count >= FoundDependenciesMinimum + 1 // Add 1 for TempPoint
                && FoundDependencies.Contains(point))
            {
                RemoveIntermediateFigureIfNecessary();
                RemoveTempPointIfNecessary();
                AddFiguresAndRestart();
                return;
            }
            base.Click(coordinates);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return null;
        }

        protected override bool CanCreateTempResults()
        {
            return FoundDependencies.Count >= 3;
        }

        protected override Type GetExpectedDependencyType()
        {
            return typeof(IPoint);
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            var polygon = Factory.CreatePolygon(Drawing, FoundDependencies);
            //if (Style != null)
            //{
            //    polygon.Style = Style;
            //}
            yield return polygon;
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
                    var segment = Factory.CreateSegment(Drawing, p1, p2);
                    yield return segment;
                }
            }
        }

        protected override IFigure CreateIntermediateFigure()
        {
            if (!FoundDependencies.All(f => f is IPoint))
            {
                return null;
            }
            if (FoundDependencies.Count == 2)
            {
                return Factory.CreateSegment(Drawing, FoundDependencies);
            }
            return null;
        }

        public override string Name
        {
            get { return "Polygon"; }
        }

        public override string HintText
        {
            get
            {
                return "Click points to construct a polygon. Click the first point again to close the polygon.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder
                .BuildIcon()
                .Polygon(
                    Factory.CreateDefaultFillBrush(),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.2, 0.4),
                    new Point(0.3, 0.8),
                    new Point(0.7, 0.8),
                    new Point(0.8, 0.4),
                    new Point(0.6, 0.2))
                .Canvas;
        }
    }
}