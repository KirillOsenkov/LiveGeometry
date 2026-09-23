using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;
using Avalonia.Media;

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

        /// <summary>
        /// A segment (anything with a length) under the cursor with no point on top of it:
        /// a click measures it instead of putting a point on it.
        /// </summary>
        IFigure FindFigureToMeasure(Point coordinates)
        {
            var underMouse = Drawing.Figures.HitTest(coordinates, f => f is ILengthProvider && !f.DependsOn(TempPoint));
            if (underMouse != null && Drawing.Figures.HitTest<IPoint>(coordinates) == null)
            {
                return underMouse;
            }

            return null;
        }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            var underMouse = FindFigureToMeasure(Coordinates(e));
            if (underMouse != null)
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

        // the hover preview tells the same story as the click: over a segment there is no
        // ghost point, the segment itself lights up and the cursor is a hand

        protected override PointPlacement FindPointPlacement(Point unconstrainedCoordinates, Point coordinates)
        {
            if (FindFigureToMeasure(unconstrainedCoordinates) != null)
            {
                return null;
            }

            return base.FindPointPlacement(unconstrainedCoordinates, coordinates);
        }

        protected override IFigure FindFigureToPick(Point unconstrainedCoordinates)
        {
            return FindFigureToMeasure(unconstrainedCoordinates) ?? base.FindFigureToPick(unconstrainedCoordinates);
        }

        protected override Avalonia.Input.Cursor GetCursor(Point coordinates)
        {
            if (FindFigureToMeasure(coordinates) != null)
            {
                return HandCursor;
            }

            return base.GetCursor(coordinates);
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