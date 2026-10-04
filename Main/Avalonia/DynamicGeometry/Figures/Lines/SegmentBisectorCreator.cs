using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace DynamicGeometry
{
    // Off the ribbon: Perpendicular does the same through the midpoint, and a click on a
    // segment here took its two ends where every other tool puts a point on it. Kept for
    // the Figure List's icon and for when it earns its place back.
    [Ignore]
    [Category(BehaviorCategories.Lines)]
    [Order(7)]
    public class SegmentBisectorCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateSegmentBisector(Drawing, FoundDependencies);
        }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            // A segment stands for its two ends only as the first click. As the second - the
            // other point, put on a segment - its ends were added to the point already
            // there: the bisector was that of the point and one end, and the click's own
            // point was left over.
            var coordinates = Coordinates(e, false, false, false);
            var underMouse = FoundDependencies.IsEmpty() ? Drawing.Figures.HitTest<Segment>(coordinates) : null;
            if (underMouse != null
                && underMouse.Dependencies.Count() == 2
                && Drawing.Figures.HitTest<IPoint>(coordinates) == null)
            {
                FoundDependencies.AddRange(underMouse.Dependencies);
            }

            base.MouseDown(sender, e);
        }

        public override string Name
        {
            get { return "Perpendicular Bisector"; }
        }

        public override string HintText
        {
            get
            {
                return "Click two points or a segment to create the perpendicular bisector.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Line(0.25, 0.75, 0.75, 0.25)
                .AccentLine(0, 0, 1, 1)
                .Point(0.25, 0.75)
                .Point(0.75, 0.25)
                .Canvas;
        }
    }
}