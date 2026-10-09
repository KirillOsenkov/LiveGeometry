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

        /// <summary>
        /// Inside an angle next to its vertex, the first click takes the angle whole, as the
        /// Angle tool does (<see cref="AngleAtVertex"/>); the later clicks are points on the sides
        /// </summary>
        protected override bool TakesAngleAtVertex()
        {
            return FoundDependencies.IsEmpty();
        }

        /// <summary>The hover shows the bisector the click would make, faint, with the angle</summary>
        protected override bool PreviewsAngleBisector()
        {
            return true;
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
                return "Click an angle vertex, then click two points on the angle sides to create an angle bisector. You can also click an angle measurement, or inside an angle next to its vertex.";
            }
        }

        /// <summary>After the vertex: which side the point is for (the bisector halves the angle under 180° whichever side comes first)</summary>
        public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
        {
            return AngleSideHint(ClickedDependencies) ?? base.ConstructionHintText(args);
        }

        /// <summary>
        /// The hint of the Angle and Angle Bisector tools after the vertex: a point on the first
        /// side, then on the second; null for any other step
        /// </summary>
        public static string AngleSideHint(int pointsClicked)
        {
            switch (pointsClicked)
            {
                case 1:
                    return "Select a point on the first side of the angle.";
                case 2:
                    return "Select a point on the second side of the angle.";
                default:
                    return null;
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