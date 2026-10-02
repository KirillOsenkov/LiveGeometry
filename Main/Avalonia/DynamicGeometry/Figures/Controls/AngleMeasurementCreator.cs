using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using Avalonia;
using Avalonia.Controls.Shapes;
using Avalonia.Media;
using M = System.Math;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Measure)]
    [Order(2)]
    public class AngleMeasurementCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPointPoint;
        }

        // The mark and the number, both made when the third point is there. (The arc used
        // to be added for good with the preview, before the third point existed: it sat in
        // the figure list ahead of a point it is built on, and a saved file came back in
        // another order.)
        protected override IEnumerable<IFigure> CreateFigures()
        {
            var sides = InsideOrder(FoundDependencies);
            yield return Factory.CreateAngleArc(Drawing, sides);
            yield return Factory.CreateAngleMeasurement(Drawing, sides);
        }

        /// <summary>
        /// The angle under 180 degrees between the two sides, whichever side was clicked
        /// first: an angle goes counterclockwise from its first side to its second, and
        /// clicked the other way round the angle of a triangle said 270 or 300 degrees.
        /// "Convert to opposite angle" gives the other one; dragged afterwards, the angle is
        /// what it has become (as the bisector's oriented sweep is).
        /// </summary>
        static IList<IFigure> InsideOrder(IList<IFigure> found)
        {
            if (found.Count == 3
                && found[0] is IPoint vertex
                && found[1] is IPoint first
                && found[2] is IPoint second
                && Math.OAngle(first.Coordinates, vertex.Coordinates, second.Coordinates) > Math.PI)
            {
                return new[] { found[0], found[2], found[1] };
            }

            return found;
        }

        /// <summary>
        /// Near a vertex, inside the mark it would get, the first click measures the angle
        /// whole (see <see cref="AngleAtVertex"/>) instead of taking a point: the hover shows
        /// it, the cursor is a hand, and no point is offered there.
        /// </summary>
        AngleAtVertex FindAngleAtVertex(Point unconstrainedCoordinates)
        {
            // the later clicks are points on the sides
            return FoundDependencies.IsEmpty() ? AngleAtVertex.Find(Drawing, unconstrainedCoordinates) : null;
        }

        protected override AngleAtVertex GetAngleToPick(MouseEventArgs e)
        {
            return FindAngleAtVertex(Coordinates(e, false, false, false));
        }

        protected override PointPlacement FindPointPlacement(Point unconstrainedCoordinates, Point coordinates)
        {
            if (FindAngleAtVertex(unconstrainedCoordinates) != null)
            {
                return null;
            }

            return base.FindPointPlacement(unconstrainedCoordinates, coordinates);
        }

        protected override Avalonia.Input.Cursor GetCursor(Point coordinates)
        {
            return FindAngleAtVertex(coordinates) != null ? HandCursor : base.GetCursor(coordinates);
        }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            // where the cursor is, as the hover preview asks: not where Shift snaps it to
            var angle = FindAngleAtVertex(Coordinates(e, false, false, false));
            if (angle != null)
            {
                FoundDependencies.AddRange(angle.Points);
                AddFiguresAndRestart();
                return;
            }

            base.MouseDown(sender, e);
        }

        public override string Name
        {
            get { return "Angle"; }
        }

        public override string HintText
        {
            get
            {
                return "Click an angle vertex, and then click two points on the angle sides to measure the angle, or click inside an angle next to its vertex.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            var builder = IconBuilder.BuildIcon();
            var size = IconBuilder.IconSize;
            var centerX = (size - 4) / (2 * size);
            var centerY = 1 - 4 / size;
            builder.Line(centerX, centerY, 0.8, 0.2)
                .Line(centerX, centerY, 1, centerY);

            // path data takes a point as the decimal separator whatever the user's culture
            var pathData = string.Format(
                CultureInfo.InvariantCulture,
                "m 0,{0} v-4 a {1},{1} 0 0 1 {2},0 v4 z m {4},-5 a 6,6 0 0 1 {3},0 z",
                size, (size - 4) / 2, size - 4, size / 2, (size - 4) / 4 - 1);
            var path = new Path
            {
                StrokeThickness = 1,
                Data = Geometry.Parse(pathData)
            };
            path.BindTheme(Shape.FillProperty, nameof(AppTheme.AngleFill));
            path.BindTheme(Shape.StrokeProperty, nameof(AppTheme.AngleOutline));
            builder.Canvas.Children.Add(path);

            var radius = ((size - 4) / 2) * 0.95;
            var radiusSmall = radius * 0.8;
            Point center = new Point(radius + 1, size - 4);
            for (double i = 0; i < 16; i++)
            {
                var angle = i * Math.PI / 15;
                builder.Line(nameof(AppTheme.ScaleMarks),
                    (center.X + radius * M.Cos(angle)) / size,
                    (center.Y - radius * M.Sin(angle)) / size,
                    (center.X + radiusSmall * M.Cos(angle)) / size,
                    (center.Y - radiusSmall * M.Sin(angle)) / size);
            }

            return builder.Canvas;
        }
    }
}