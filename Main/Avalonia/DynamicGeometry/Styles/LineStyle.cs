using Avalonia.Media;
using Avalonia.Controls.Shapes;

namespace DynamicGeometry
{
    [StyleFor(typeof(ILinearFigure))]
    public class LineStyle : FigureStyle
    {
        public override FrameworkElement GetSampleGlyph()
        {
            var line = Factory.CreateLineShape();
            line.X1 = 0;
            line.X2 = 20;
            line.Y1 = 20;
            line.Y2 = 0;
            line.Apply(this.GetWpfStyle());
            OnApplied(null, line);
            line.Tag = this;
            return line;
        }

        Color mColor = Color.FromArgb(100, 0, 0, 0);
        [PropertyGridVisible]
        [PropertyGridName("Stroke color")]
        public Color Color
        {
            get
            {
                return mColor;
            }
            set
            {
                mColor = value;
                OnPropertyChanged("Color");
            }
        }

        double strokeWidth = 1;
        [PropertyGridName("Stroke width")]
        [PropertyGridVisible]
        [Domain(0.1, 50)]
        public virtual double StrokeWidth
        {
            get
            {
                return (double)strokeWidth;
            }
            set
            {
                strokeWidth = value;
                OnPropertyChanged("StrokeWidth");
            }
        }

        LineDash dash = LineDash.Solid;
        [PropertyGridVisible]
        [PropertyGridName("Dash")]
        public LineDash Dash
        {
            get
            {
                return dash;
            }
            set
            {
                dash = value;
                OnPropertyChanged("Dash");
            }
        }

        protected override void ApplyToWpfStyle(Style existingStyle, IFigure figure)
        {
            base.ApplyToWpfStyle(existingStyle, figure);
            double width = strokeWidth;
            if (figure != null && figure.Selected && Settings.ChangeLineAppearanceWhenSelected)
            {
                width += 3;
            }

            var strokeSetter = new Setter(Shape.StrokeProperty, new SolidColorBrush(Color));
            existingStyle.Setters.Add(strokeSetter);

            var widthSetter = new Setter(Shape.StrokeThicknessProperty, width);
            existingStyle.Setters.Add(widthSetter);

        }

        /// <summary>
        /// The dash pattern goes on here and not through a setter: it depends on the width the
        /// shape ended up with (a selected figure is drawn thicker).
        /// </summary>
        public override void OnApplied(IFigure figure, FrameworkElement element)
        {
            var shape = element as Shape;
            if (shape != null)
            {
                shape.StrokeDashArray = LineDashes.GetDashArray(Dash, shape.StrokeThickness);
            }
        }
    }
}