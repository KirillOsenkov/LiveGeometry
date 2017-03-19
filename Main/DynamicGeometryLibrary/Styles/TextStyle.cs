using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DynamicGeometry
{
    [StyleFor(typeof(LabelBase))]
    public class TextStyle : FigureStyle
    {
        public override FrameworkElement GetSampleGlyph()
        {
            var label = Factory.CreateLabelShape();
            label.Text = "Text";
            label.Apply(this.GetWpfStyle());
            label.Tag = this;
            return label;
        }

        protected double fontSize = 10.0;
        [PropertyGridName("Font size")]
        [PropertyGridVisible]
        [Domain(3, 150)]
        public double FontSize
        {
            get
            {
                return fontSize;
            }
            set
            {
                fontSize = value;
                OnPropertyChanged("FontSize");
            }
        }

        Color mColor = Colors.Black;
        [PropertyGridVisible]
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

        FontFamily fontFamily = new FontFamily("Arial");
        [PropertyGridVisible]
        [PropertyGridName("Font family")]
        public FontFamily FontFamily
        {
            get
            {
                return fontFamily;
            }
            set
            {
                fontFamily = value;
                OnPropertyChanged("FontFamily");
            }
        }

        bool bold = false;
        [PropertyGridVisible]
        public bool Bold
        {
            get
            {
                return bold;
            }
            set
            {
                bold = value;
                OnPropertyChanged("Bold");
            }
        }

        bool italic = false;
        [PropertyGridVisible]
        public bool Italic
        {
            get
            {
                return italic;
            }
            set
            {
                italic = value;
                OnPropertyChanged("Italic");
            }
        }

        bool underline = false;
        [PropertyGridVisible]
        public virtual bool Underline
        {
            get
            {
                return underline;
            }
            set
            {
                underline = value;
                OnPropertyChanged("Underline");
            }
        }

        protected override void ApplyToWpfStyle(Style existingStyle, IFigure figure)
        {
            base.ApplyToWpfStyle(existingStyle, figure);
            existingStyle.Setters.Add(new Setter(TextBlock.FontSizeProperty, fontSize));
            existingStyle.Setters.Add(new Setter(TextBlock.ForegroundProperty, new SolidColorBrush(Color)));
            if (FontFamily != null && !FontFamily.Source.IsEmpty())
            {
                existingStyle.Setters.Add(new Setter(TextBlock.FontFamilyProperty, FontFamily));
            }
            existingStyle.Setters.Add(new Setter(TextBlock.FontStyleProperty, Italic ? FontStyles.Italic : FontStyles.Normal));
            existingStyle.Setters.Add(new Setter(TextBlock.FontWeightProperty, Bold ? FontWeights.Bold : FontWeights.Normal));
            if (Underline)
            {
                existingStyle.Setters.Add(new Setter(TextBlock.TextDecorationsProperty, TextDecorations.Underline));
            }
            //existingStyle.Setters.Add(new Setter(TextBlock.HorizontalAlignmentProperty, Alignment));
        }
    }
}