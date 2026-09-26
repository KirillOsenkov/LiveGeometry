using System;
using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Circles)]
    [Order(2)]
    public class CircleByRadiusCreator : FigureCreator
    {
        public CircleByRadiusCreator()
        {
            CanReuseDependency = true;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateCircleByRadius(Drawing, FoundDependencies);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPointPoint;
        }

        protected override IFigure CreateIntermediateFigure()
        {
            if (FoundDependencies.Count == 2
                && FoundDependencies[0] is IPoint
                && FoundDependencies[1] is IPoint)
            {
                return Factory.CreateSegment(Drawing, FoundDependencies);
            }
            return null;
        }

        // Anything with a length can stand for the two radius points: the first click on one
        // with no point on top of it takes it, and the next click is the center. A segment, a
        // vector or a distance measurement is unwrapped into its two points (see
        // FindRadiusEnds); anything else with a length is the radius itself.

        bool RadiusIsAFigure
        {
            get { return FoundDependencies.Count > 0 && !(FoundDependencies[0] is IPoint); }
        }

        protected override IFigure FindFigureInsteadOfPoint(Point unconstrainedCoordinates)
        {
            if (FoundDependencies.Count > 0)
            {
                return null;
            }

            var underMouse = Drawing.Figures.HitTest(
                unconstrainedCoordinates,
                f => f is ILengthProvider && f.Visible && f.IsHitTestVisible);
            if (underMouse != null && Drawing.Figures.HitTest<IPoint>(unconstrainedCoordinates) == null)
            {
                return underMouse;
            }

            return null;
        }

        /// <summary>
        /// The two points a figure with a length is built on, when it is: a segment or a vector
        /// (two point dependencies), or a distance measurement of two points or of such a
        /// figure. The circle then depends on the points, as if they had been clicked, and
        /// outlives the figure. Null for a polyline, an arc or a label with an expression: those
        /// are the radius themselves.
        /// </summary>
        static IList<IPoint> FindRadiusEnds(IFigure figure)
        {
            if (figure is DistanceMeasurement measurement && measurement.Dependencies[0] is ILengthProvider measured)
            {
                return FindRadiusEnds(measured);
            }

            if (figure.Dependencies.Count == 2
                && figure.Dependencies[0] is IPoint first
                && figure.Dependencies[1] is IPoint second)
            {
                return new[] { first, second };
            }

            return null;
        }

        protected override Type GetExpectedDependencyType()
        {
            if (TempPoint == null && RadiusIsAFigure && FoundDependencies.Count == 2)
            {
                return null;
            }

            return base.GetExpectedDependencyType();
        }

        protected override void AddDependency(Point coordinates)
        {
            var radius = FindFigureInsteadOfPoint(ClickedUnconstrainedCoordinates);
            if (radius == null)
            {
                base.AddDependency(coordinates);
                return;
            }

            StartConstruction();
            var ends = FindRadiusEnds(radius);
            if (ends != null)
            {
                FoundDependencies.AddRange(ends);
            }
            else
            {
                FoundDependencies.Add(radius);
            }

            // the center follows the cursor, with the circle already around it
            CreateTempPoint(coordinates);
            CreateTempResults();
            AdvertiseNextDependency();
            Drawing.Figures.CheckConsistency();
        }

        public override string Name
        {
            get
            {
                return "By Radius";
            }
        }

        public override string HintText
        {
            get
            {
                return "Click two points (the ends of a radius), a segment or a distance, then click the circle center.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            const double r = 0.4;
            return IconBuilder.BuildIcon()
                .Circle(r, r, r)
                .Line(0.5, 0.9, 0.5 + r, 0.9)
                .Point(r, r)
                .Point(0.5, 0.9)
                .Point(0.5 + r, 0.9)
                .Canvas;
        }
    }
}