using Avalonia.Media;

namespace DynamicGeometry
{
    /// <summary>
    /// An axis of the coordinate grid: a line drawn as a vector's arrow, the shaft across
    /// the window and the head where the axis leaves it, in the direction of the axis
    /// (right, up). The line itself is invisible, as a vector's segment is.
    /// </summary>
    public class Axis : CompositeFigure
    {
        public LineTwoPoints Line { get; set; }
        public Arrow Arrow { get; set; }

        public Axis()
        {
            Line = new LineTwoPoints();
            Line.Style = new LineStyle() { StrokeWidth = 0, Color = Colors.Transparent };
            Line.Layer = ZOrder.Axes;
            Arrow = new Arrow();
            Arrow.Layer = ZOrder.Axes;
            Arrow.Dependencies.Add(Line);
            Children.Add(Line, Arrow);
        }

        public override IFigure HitTest(Avalonia.Point point, System.Predicate<IFigure> filter)
        {
            return null;
        }

        /// <summary>The theme's axis color, as on screen now</summary>
        public static Color Color
        {
            get
            {
                return AppTheme.Current.Axis;
            }
        }
    }
}
