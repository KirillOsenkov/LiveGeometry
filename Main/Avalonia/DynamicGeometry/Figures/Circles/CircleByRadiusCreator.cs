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

        /// <summary>
        /// The second end of the radius on the first is a radius of 0; the center may be
        /// either end (the circles about A and about B of an equilateral triangle)
        /// </summary>
        protected override bool IsDegenerateRepeat(IFigure figure, IList<IFigure> found)
        {
            return found.Count == 1 && found[0] == figure;
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
            if (FoundDependencies.Count > 0 || slider.Exists)
            {
                return null;
            }

            var underMouse = Drawing.Figures.HitTest(
                unconstrainedCoordinates,
                f => f.GivesLength() && f.Visible && f.IsHitTestVisible);
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
            TakeRadius(radius, coordinates);
        }

        /// <summary>The radius is this figure's length (or the distance between its two points); the center is next</summary>
        void TakeRadius(IFigure radius, Point coordinates)
        {
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
            Drawing.Figures.CheckConsistencyInDebug();
        }

        #region A slider for the radius

        // A first click on empty paper, where a free point would appear, starts a slider
        // instead, placed as the Slider tool places one: the click is its anchor, the knob
        // follows the cursor, the second click (or the release of a press and drag) says
        // where the knob starts. The slider is then the radius, and the center is next.
        // Typed coordinates (Point by coordinates) still give a point: they come in through
        // AddDependency, not through Click.

        readonly PendingSlider slider = new PendingSlider();

        bool StartsSlider(Point coordinates)
        {
            if (FoundDependencies.Count > 0)
            {
                return false;
            }

            var placement = FindPointPlacement(ClickedUnconstrainedCoordinates, coordinates);
            return placement != null && placement.Kind == PointPlacementKind.Free;
        }

        protected override void Click(Point coordinates)
        {
            if (slider.Exists)
            {
                // (not where it began: that is the second click of a double click, and a
                // slider of length 0 would be a circle of radius 0)
                if (slider.IsDragged(coordinates))
                {
                    FinishSlider(coordinates);
                }
            }
            else if (StartsSlider(coordinates))
            {
                StartConstruction();
                ConstructionComplete = false;
                slider.Start(Drawing, coordinates);
                Drawing.RaiseStatusNotification("Click where the slider ends: its length is the radius.");
            }
            else
            {
                base.Click(coordinates);
            }
        }

        void FinishSlider(Point coordinates)
        {
            // recorded in the construction's transaction: undo takes the circle and its slider
            var radius = slider.Finish(coordinates);
            Actions.Add(Drawing, radius);
            TakeRadius(radius, coordinates);
        }

        public override void MouseMove(object sender, MouseEventArgs e)
        {
            base.MouseMove(sender, e);
            slider.Follow(Coordinates(e));
        }

        public override void MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!slider.Exists)
            {
                base.MouseUp(sender, e);
                return;
            }

            // not the base: once the slider is there the center follows the cursor, and the
            // base would take this release for the click that places it
            IsMouseButtonDown = false;
            var coordinates = Coordinates(e);
            if (slider.IsDragged(coordinates))
            {
                FinishSlider(coordinates);
            }
        }

        /// <summary>No ghost point while the knob follows the cursor</summary>
        protected override PointPlacement FindPointPlacement(Point unconstrainedCoordinates, Point coordinates)
        {
            return slider.Exists ? null : base.FindPointPlacement(unconstrainedCoordinates, coordinates);
        }

        protected override Avalonia.Input.Cursor GetCursor(Point coordinates)
        {
            return slider.Exists ? ArrowCursor : base.GetCursor(coordinates);
        }

        public override bool IsInInitialState
        {
            get { return !slider.Exists && base.IsInInitialState; }
        }

        public override void Stopping()
        {
            slider.Cancel();
            base.Stopping();
        }

        #endregion

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
                return "Click two points (the ends of a radius), a segment, a distance or a slider - or empty paper, to make a slider for the radius. Then click the circle center.";
            }
        }

        public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
        {
            // the point that follows the cursor is among the found ones
            bool radiusIsKnown = RadiusIsAFigure || FoundDependencies.Count == 3;
            return radiusIsKnown ? "Click the circle center." : "Click the other end of the radius.";
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