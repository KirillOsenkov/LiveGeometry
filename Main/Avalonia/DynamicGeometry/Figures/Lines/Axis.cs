using Avalonia.Media;

namespace DynamicGeometry
{
    public class Axis : CompositeFigure
    {
        public LineTwoPoints Line { get; set; }
        public Arrow Arrow { get; set; }

        public Axis()
        {
            Line = new LineTwoPoints();
            Line.SetZIndex(ZOrder.Axes);

            //Arrow = new Arrow();
            //Arrow.Dependencies.Add(Line);
            //Children.Add(Line, Arrow);
            Children.Add(Line);
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
