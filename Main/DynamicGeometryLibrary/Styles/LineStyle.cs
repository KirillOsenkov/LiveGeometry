using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;

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
        [PropertyGridName("Stroke Color")]
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

        DoubleCollection strokeDashArray;
        [Ignore]
        public DoubleCollection StrokeDashArray
        {
            get
            {
                return strokeDashArray;
            }
            set
            {
                strokeDashArray = value;
                OnPropertyChanged("StrokeDashArray");
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

            // TODO: There is a known bug in Silverlight 2/3/4 where setting
            // StrokeDashArray via a Style results in an ArgumentException
            // Apply workaround in OnApplied()
#if !SILVERLIGHT
            if (!StrokeDashArray.IsEmpty())
            {
                var strokeDashArraySetter = new Setter(Shape.StrokeDashArrayProperty, StrokeDashArray);
                existingStyle.Setters.Add(strokeDashArraySetter);
            }
#endif
        }

        public override void OnApplied(IFigure figure, FrameworkElement element)
        {
            var line = element as Line;
            
            if (StrokeDashArray.IsEmpty())
            {
                if (line != null)
                {
                    line.StrokeDashArray = null;
                }
                return;
            }

            if (line != null)
            {
                var collection = new DoubleCollection();
                collection.AddRange(this.StrokeDashArray);
                line.StrokeDashArray = collection;
            }
        }
    }
}