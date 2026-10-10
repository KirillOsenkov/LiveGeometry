using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;

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

        /// <summary>
        /// An angle's mark or number, or an arc, a sector or a segment, under the cursor
        /// with no point on top of it: the first click takes its three points and its sweep
        /// (<see cref="AnglePoints"/>). Only as the first click (see SegmentBisectorCreator):
        /// after a vertex was clicked, a click on an angle's mark gave a bisector of five
        /// points, which doesn't exist, and a stray point.
        /// </summary>
        protected override IReadOnlyList<IFigure> FindFiguresInsteadOfPoint(Point unconstrainedCoordinates)
        {
            return FindAngleFigures(this, unconstrainedCoordinates, AnglePoints.Takes);
        }

        /// <summary>
        /// The figures a click takes an angle from, for this tool and the Angle tool: those
        /// the filter takes, in view, and none while a point is under the cursor or a step
        /// has been made
        /// </summary>
        public static IReadOnlyList<IFigure> FindAngleFigures(FigureCreator creator, Point unconstrainedCoordinates, System.Func<IFigure, bool> takes)
        {
            var drawing = creator.Drawing;
            if (drawing == null || !creator.IsInInitialState || drawing.Figures.HitTest<IPoint>(unconstrainedCoordinates) != null)
            {
                return System.Array.Empty<IFigure>();
            }

            return drawing.Figures.HitTestAll(
                unconstrainedCoordinates,
                f => takes(f) && f.Visible && f.IsHitTestVisible);
        }

        AngleSweep? clickedSweep;

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            // where the cursor is, as the hover preview asks: not where Shift snaps it to
            var angle = AnglePoints.From(FindFigureInsteadOfPoint(Coordinates(e, false, false, false)));
            if (angle != null)
            {
                StartConstruction();
                FoundDependencies.AddRange(angle.Points);
                clickedSweep = angle.Sweep;
                AddFiguresAndRestart();
                return;
            }

            base.MouseDown(sender, e);
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

        /// <summary>The bisector; of the angle a clicked figure shows, when it was one (its sweep copied once)</summary>
        protected override IEnumerable<IFigure> CreateFigures()
        {
            var result = Factory.CreateAngleBisector(Drawing, FoundDependencies);
            if (clickedSweep != null)
            {
                result.Sweep = clickedSweep.Value;
                clickedSweep = null;
            }

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
                return "Click an angle vertex, then click two points on the angle sides to create an angle bisector. You can also click an angle measurement, an arc, or inside an angle next to its vertex.";
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