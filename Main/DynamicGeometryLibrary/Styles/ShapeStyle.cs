using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    [StyleFor(typeof(IShapeWithInterior))]
    [StyleFor(typeof(Bezier))]
    public class ShapeStyle : LineStyle
    {
        public override FrameworkElement GetSampleGlyph()
        {
            var polygon = Factory.CreatePolygonShape();
            polygon.Points = new PointCollection() 
            {
                new Point(0, 20),
                new Point(10, 0),
                new Point(20, 20)
            };
            polygon.Apply(this.GetWpfStyle());
            polygon.Tag = this;
            return polygon;
        }

        Brush mFill = new SolidColorBrush(Colors.Yellow);
        [PropertyGridVisible]
        public Brush Fill
        {
            get
            {
                return mFill;
            }
            set
            {
                mFill = value;
                OnPropertyChanged("Fill");
            }
        }

        bool mIsFilled = true;
        [PropertyGridVisible]
        public bool IsFilled
        {
            get
            {
                return mIsFilled;
            }
            set
            {
                mIsFilled = value;
                OnPropertyChanged("IsFilled");
            }
        }

        protected override void ApplyToWpfStyle(Style existingStyle, IFigure figure)
        {
            base.ApplyToWpfStyle(existingStyle, figure);
            var brush = Fill;

            brush = ModifyFillBrushIfSelected(figure, brush, this);

            if (!IsFilled)
            {
                brush = null;
            }
            var fillSetter = new Setter(Shape.FillProperty, brush);
            existingStyle.Setters.Add(fillSetter);
            var miterLimitSetter = new Setter(Shape.StrokeMiterLimitProperty, 1.0);
            existingStyle.Setters.Add(miterLimitSetter);
        }

        public static Brush ModifyFillBrushIfSelected(IFigure figure, Brush brush, FigureStyle style)
        {
            // Below is my suggestion for the fill appearance when a shape is selected. - David
            var brushAsSolidColor = brush as SolidColorBrush;
            if (figure != null && figure.Selected && !(style is PointStyle) && brushAsSolidColor != null)
            {
                // brush.Opacity = 0.2; // The previous method of showing a figure is selected.

                // The color of the stripes in the gradient is made by shifting the Fill color toward a gray value.
                var b = new LinearGradientBrush();
                var shapeWidth = 200;   // Arbitrary value.  Using the with of the figure would be better.
                b.StartPoint = new Point(0.0, 0.0);
                b.EndPoint = new Point(1.0, 0.0);
                double gap = 6;    // Actually half the gap between stripes in physical coordinates.
                double s = gap / shapeWidth;
                byte gray = 200;    // The gray value to shift the Fill color toward.
                double similarity = .65;    // How similar are the Fill color and the gray value (1 = similar, 0 = not).
                byte alpha = ((int)brushAsSolidColor.Color.A < 128) ? (byte)128 : brushAsSolidColor.Color.A;  // We need some opacity.
                for (double i = 0; i + s + s < 1; i += 2 * s)
                {
                    b.GradientStops.Add(new GradientStop() { Color = brushAsSolidColor.Color, Offset = i + .7 * s });
                    b.GradientStops.Add(new GradientStop()
                    {
                        Color = Color.FromArgb(alpha, (byte)(gray + (brushAsSolidColor.Color.R - gray) * similarity),
                                                      (byte)(gray + (brushAsSolidColor.Color.G - gray) * similarity),
                                                      (byte)(gray + (brushAsSolidColor.Color.B - gray) * similarity)),
                        Offset = i + s
                    });
                    b.GradientStops.Add(new GradientStop() { Color = brushAsSolidColor.Color, Offset = i + s + .3 * s });
                }
                brush = b;
            }
            // End of suggestion.
            return brush;
        }
    }
}