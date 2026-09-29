using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry
{
    public class StatusBar : WrapPanel
    {
        public TextBlock TextBlock { get; set; }
        public Border border = new Border();

        public StatusBar()
        {
            HorizontalAlignment = HorizontalAlignment.Left;
            VerticalAlignment = VerticalAlignment.Bottom;
            Margin = new Thickness(8);

            border.Padding = new Thickness(10, 5, 10, 5);
            border.CornerRadius = new CornerRadius(6);
            border.BindTheme(Border.BackgroundProperty, nameof(AppTheme.HintBackground));
            border.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.HintBorder));
            border.BorderThickness = new Thickness(1);
            border.BoxShadow = new BoxShadows(new BoxShadow()
            {
                OffsetY = 1,
                Blur = 4,
                Color = Color.FromArgb(40, 0, 0, 0)
            });
            border.PointerPressed += border_MouseLeftButtonDown;

            TextBlock = new TextBlock()
            {
                FontSize = 12,
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 640
            };
            TextBlock.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
            border.Child = TextBlock;

            this.Children.Add(border);
        }

        public string Text
        {
            get
            {
                return TextBlock.Text;
            }
            set
            {
                TextBlock.Text = value;
            }
        }

#if !SILVERLIGHT
        new 
#endif
        public bool IsVisible
        {
            get
            {
                return this.Visibility == Visibility.Visible;
            }
            set
            {
                this.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        void border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            IsVisible = false;
        }
    }
}
