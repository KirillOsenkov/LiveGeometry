using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Points)]
    [Order(2)]
    public class MidpointCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            // one midpoint per pair of points is enough
            if (PointPlacement.FindExistingMidpoint(FoundDependencies[0], FoundDependencies[1]) != null)
            {
                yield break;
            }

            MidPoint result = Factory.CreateMidPoint(Drawing, FoundDependencies);
            yield return result;
        }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            var underMouse = FindSegmentToBisect(e);
            if (underMouse != null)
            {
                if (HasMidpointAlready(underMouse))
                {
                    return;
                }

                FoundDependencies.AddRange(underMouse.Dependencies);
            }

            base.MouseDown(sender, e);
        }

        static bool HasMidpointAlready(Segment segment)
        {
            return PointPlacement.FindExistingMidpoint(segment.Dependencies[0], segment.Dependencies[1]) != null;
        }

        /// <summary>
        /// A click on a segment (not on a point of it) takes both its endpoints at once
        /// </summary>
        Segment FindSegmentToBisect(MouseEventArgs e)
        {
            if (!FoundDependencies.IsEmpty())
            {
                return null;
            }

            var coordinates = Coordinates(e);
            var segment = Drawing.Figures.HitTest<Segment>(coordinates);
            if (segment == null
                || !PointPlacement.HasMidpoint(segment)
                || Drawing.Figures.HitTest<IPoint>(coordinates) != null)
            {
                return null;
            }

            return segment;
        }

        protected override PointPlacement GetClickPreview(MouseEventArgs e)
        {
            var segment = FindSegmentToBisect(e);
            if (segment != null)
            {
                return HasMidpointAlready(segment) ? null : PointPlacement.Midpoint(segment);
            }

            return base.GetClickPreview(e);
        }

        public override string Name
        {
            get { return "Midpoint"; }
        }

        public override string HintText
        {
            get
            {
                return "Click two points (or a segment) to construct a midpoint.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Point(0.25, 0.75)
                .DependentPoint(0.5, 0.5)
                .Point(0.75, 0.25)
                .Canvas;
        }
    }
}