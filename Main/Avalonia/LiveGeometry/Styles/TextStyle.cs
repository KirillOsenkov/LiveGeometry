using Avalonia.Controls;
using Avalonia.Media;

namespace LiveGeometry;

[StyleFor(typeof(LabelBase))]
[StyleFor(typeof(ShowHideControl))]
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
        get => fontSize;
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
        get => mColor;
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
        get => fontFamily;
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
        get => bold;
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
        get => italic;
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
        get => underline;
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
        if (FontFamily != null && !FontFamily.Name.IsEmpty())
        {
            // A drawing names fonts of the machine it was made on (Segoe UI, Arial); the
            // browser has none of them, and what a missing font falls back to by itself
            // there is Noto Mono - and so are FontFamily.Default and the font manager's
            // default. The app's own font (embedded Inter) is what the theme puts on the
            // window and text inherits, so a font that isn't there is simply not set.
            bool isInstalled = FontManager.Current.TryGetGlyphTypeface(new Typeface(FontFamily), out var resolved)
                && string.Equals(resolved.FamilyName, FontFamily.Name, System.StringComparison.OrdinalIgnoreCase);
            if (isInstalled)
            {
                existingStyle.Setters.Add(new Setter(TextBlock.FontFamilyProperty, FontFamily));
            }
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