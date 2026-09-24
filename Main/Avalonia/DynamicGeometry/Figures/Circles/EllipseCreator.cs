using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;

namespace DynamicGeometry
{
    /// <summary>
    /// Center, the end of the long axis, the end of the short axis. A free third click lands
    /// on the short axis: once the first two points are known the tool adds a hidden segment
    /// (the long semi-axis) and a hidden perpendicular to it through the center, and the third
    /// point is put on that line, so it sits on the ellipse and slides along the axis when
    /// dragged. A click on an existing point, or on another figure, takes that instead and the
    /// two helpers go away again.
    /// </summary>
    [Category(BehaviorCategories.Circles)]
    [Order(3)]
    public class EllipseCreator : FigureCreator
    {
        Segment longAxis;
        PerpendicularLine shortAxis;

        public override void Started()
        {
            base.Started();
            longAxis = null;
            shortAxis = null;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateEllipse(Drawing, FoundDependencies);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPointPoint;
        }

        protected override IFigure CreateIntermediateFigure()
        {
            // the long semi-axis, while its end is being chosen
            if (FoundDependencies.Count == 2
                && FoundDependencies[0] is IPoint
                && FoundDependencies[1] is IPoint)
            {
                return Factory.CreateSegment(Drawing, FoundDependencies);
            }

            return null;
        }

        protected override void CreateTempResults()
        {
            // The preview ellipse is made once the center and the long axis are known and the
            // third point is being chosen; that is when the short axis appears.
            var center = FoundDependencies[0] as IPoint;
            var longAxisEnd = FoundDependencies[1] as IPoint;
            if (shortAxis == null
                && center != null
                && longAxisEnd != null
                && center.Coordinates != longAxisEnd.Coordinates)
            {
                longAxis = Factory.CreateSegment(Drawing, center, longAxisEnd);
                longAxis.Visible = false;
                shortAxis = Factory.CreatePerpendicularLine(Drawing, new IFigure[] { longAxis, center });
                shortAxis.Visible = false;
                Actions.Add(Drawing, longAxis);
                Actions.Add(Drawing, shortAxis);
            }

            base.CreateTempResults();
        }

        protected override PointPlacement FindPointPlacement(Point unconstrainedCoordinates, Point coordinates)
        {
            var placement = base.FindPointPlacement(unconstrainedCoordinates, coordinates);
            if (shortAxis != null && placement != null && placement.Kind == PointPlacementKind.Free)
            {
                return PointPlacement.OnFigure(shortAxis, coordinates);
            }

            return placement;
        }

        protected override void AddFiguresAndRestart()
        {
            bool onShortAxis = FoundDependencies.Count == 3
                && FoundDependencies[2] is PointOnFigure point
                && point.Dependencies[0] == shortAxis;
            if (shortAxis != null && !onShortAxis)
            {
                Actions.Remove(shortAxis);
                Actions.Remove(longAxis);
            }

            base.AddFiguresAndRestart();
        }

        public override string Name
        {
            get
            {
                return "Ellipse";
            }
        }

        public override string HintText
        {
            get
            {
                return "Click the center, then the end of the long axis, then the end of the short axis.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Ellipse(0.5, 0.5, 0.5, 0.3)
                .Point(0.5, 0.5)
                .Point(1.0, 0.5)
                .Point(0.5, 0.2)
                .Canvas;
        }
    }
}