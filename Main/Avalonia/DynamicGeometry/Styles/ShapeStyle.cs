using Avalonia;
using Avalonia.Media;
using Avalonia.Controls.Shapes;

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
            Resolve().OnApplied(null, polygon);
            polygon.Tag = this;
            return polygon;
        }

        Brush mFill = new SolidColorBrush(Color.FromArgb(100, 255, 255, 200));
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
        [PropertyGridName("Filled")]
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

        /// <summary>
        /// How much of the stroke's color the fill of <see cref="WithStrokeOf"/> has: a hint,
        /// as the outline styles had before the palette
        /// </summary>
        const byte FillAlphaOfStroke = 28;

        /// <summary>
        /// A shape style with the line style's stroke under every theme, and a fill in the
        /// stroke's color, faint, to switch on: not filled yet, so that it looks as the line did
        /// </summary>
        public static ShapeStyle WithStrokeOf(LineStyle line)
        {
            var result = new ShapeStyle()
            {
                Color = line.Color,
                StrokeWidth = line.StrokeWidth,
                Dash = line.Dash,
                Fill = new SolidColorBrush(AppTheme.WithAlpha(line.Color, FillAlphaOfStroke)),
                IsFilled = false
            };
            foreach (var theme in line.Overrides)
            {
                foreach (var pair in theme.Value)
                {
                    result.SetOverride(theme.Key, pair.Key, pair.Value);
                    if (pair.Key == nameof(Color) && pair.Value is Color color)
                    {
                        result.SetOverride(theme.Key, nameof(Fill), new SolidColorBrush(AppTheme.WithAlpha(color, FillAlphaOfStroke)));
                    }
                }
            }

            return result;
        }

        protected override void ApplyToWpfStyle(Style existingStyle, IFigure figure)
        {
            base.ApplyToWpfStyle(existingStyle, figure);
            var brush = Fill;
            if (!IsFilled)
            {
                brush = null;
            }
            var fillSetter = new Setter(Shape.FillProperty, brush);
            existingStyle.Setters.Add(fillSetter);
            var miterLimitSetter = new Setter(Shape.StrokeMiterLimitProperty, 1.0);
            existingStyle.Setters.Add(miterLimitSetter);
        }
    }
}