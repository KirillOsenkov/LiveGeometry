using System.Collections.Generic;
using System.ComponentModel;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Lines)]
    [Order(8)]
    public class AngleBisectorCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPointPoint;
        }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            var angle = FindAngle(e);
            if (angle != null)
            {
                FoundDependencies.AddRange(angle.Dependencies);
            }

            base.MouseDown(sender, e);
        }

        /// <summary>
        /// An angle stands for its three points only as the first click (see
        /// SegmentBisectorCreator): after a vertex was clicked, a click on an angle's
        /// mark gave a bisector of five points, which doesn't exist, and a stray point.
        /// </summary>
        IFigure FindAngle(MouseEventArgs e)
        {
            var underMouse = FoundDependencies.IsEmpty() ? Drawing.Figures.HitTest(Coordinates(e, false, false, false)) : null;
            if (underMouse != null
                && (underMouse is AngleArc || underMouse is AngleMeasurement)
                && underMouse.Dependencies.Count == 3)
            {
                return underMouse;
            }

            return null;
        }

        /// <summary>A click on an angle takes its points, whatever else is there: nothing to choose</summary>
        protected override IReadOnlyList<object> FindClickOptions(MouseEventArgs e)
        {
            return FindAngle(e) != null ? System.Array.Empty<object>() : base.FindClickOptions(e);
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            var result = Factory.CreateAngleBisector(Drawing, FoundDependencies);
            yield return result;
        }

        public override string Name
        {
            get { return "Angle Bisector"; }
        }

        public override string HintText
        {
            get
            {
                return "Click an angle vertex, then click two points on the angle sides to create an angle bisector. You can also click an angle measurement.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            const double a = 0.9, b = 0.1;
            var builder = IconBuilder.BuildIcon()
                .Line(a, a, b, a)
                .Line(b, a, b, b)
                .AccentLine(a, b, b, a)
                .Arc(b, a, 0.4, a, b, 0.6)
                .Point(a, a)
                .Point(b, a)
                .Point(b, b);

            return builder.Canvas;
        }
    }
}