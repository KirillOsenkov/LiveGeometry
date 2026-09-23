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

        // A segment (anything with a length) can stand for the two radius points: the first
        // click on one with no point on top of it takes it, and the next click is the center.

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

            var underMouse = Drawing.Figures.HitTest(unconstrainedCoordinates, f => f is ILengthProvider);
            if (underMouse != null && Drawing.Figures.HitTest<IPoint>(unconstrainedCoordinates) == null)
            {
                return underMouse;
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

            Drawing.RaiseConstructionStepStarted();
            FoundDependencies.Add(radius);

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
                return "Click two points (the ends of a radius) or a segment, then click the circle center.";
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