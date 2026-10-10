using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Circles)]
    [Order(4)]
    public class CircleArcCreator : FigureCreator
    {
        public CircleArcCreator()
        {
            CanReuseDependency = true;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateArc(Drawing, FoundDependencies);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPointPoint;
        }

        /// <summary>
        /// Center, start and end are three points: the start on the center is a radius of 0,
        /// the end on the start or the center no arc (a double click made them)
        /// </summary>
        protected override bool IsDegenerateRepeat(IFigure figure, IList<IFigure> found)
        {
            return true;
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

        /// <summary>
        /// The arc just made can become a sector or a circular segment in the side panel
        /// (<see cref="ArcPanel"/>): the same three clicks make all three, so the shapes have
        /// no tool of their own
        /// </summary>
        protected override void ShowCreatedFigure(IList<IFigure> figures)
        {
            var arc = figures.OfType<EllipseArcBase>().FirstOrDefault();
            if (arc != null)
            {
                ArcPanel.Show(arc);
            }
        }

        public override string Name
        {
            get
            {
                return "Circular Arc";
            }
        }

        public override string HintText
        {
            get
            {
                return "Click (and release) the center point and then two points (start and end of the arc, counterclockwise).";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            const double r = 0.4;
            return IconBuilder.BuildIcon()
                .Arc(0.5, 0.5, 0.5 + r, 0.5, 0.5, 0.5 - r)
                .Point(0.5, 0.5 - r)
                .Point(0.5, 0.5)
                .Point(0.5 + r, 0.5)
                .Canvas;
        }
    }

    [Category(BehaviorCategories.Circles)]
    [Order(5)]
    public class EllipseArcCreator : CircleArcCreator
    {
        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateEllipseArc(Drawing, FoundDependencies);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPointPointPointPoint;
        }

        /// <summary>
        /// Center and the ends of the two axes are three points; the arc may begin or end at
        /// the end of an axis, but not at the center, and it doesn't end where it begins
        /// </summary>
        protected override bool IsDegenerateRepeat(IFigure figure, IList<IFigure> found)
        {
            return found.Count < 3 || figure == found[0] || found.Count == 4 && figure == found[3];
        }

        protected override IFigure CreateIntermediateFigure()
        {
            if (FoundDependencies.Count == 2
                && FoundDependencies[0] is IPoint
                && FoundDependencies[1] is IPoint)
            {
                return Factory.CreateSegment(Drawing, FoundDependencies);
            }
            else if (FoundDependencies.Count == 3
                && FoundDependencies[0] is IPoint
                && FoundDependencies[1] is IPoint
                && FoundDependencies[2] is IPoint)
            {
                return Factory.CreateEllipse(Drawing, FoundDependencies);
            }
            return null;
        }

        public override string Name
        {
            get
            {
                return "Elliptical Arc";
            }
        }

        public override string HintText
        {
            get
            {
                return "Click on 5 points in order to determine the center, the semi-major axis, the semi-minor axis, the begin angle, and the end angle.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            const double a = 0.4;   // semi-major
            const double b = 0.8;   // semi-minor
            Size s = new Size(a, b);
            return IconBuilder.BuildIcon()
                .EllipseArc(0.5, 0.9, 0.5 + a, 0.9, 0.5, 0.9 - b, s, 0)
                .Point(0.5, 0.9)
                .Point(0.5 + a, 0.9)
                .Point(0.5, 0.9 - b)
                .Canvas;
        }
    }
}
