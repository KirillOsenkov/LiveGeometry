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