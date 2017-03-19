using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public class ButtonGrid : Grid
    {
        public ButtonGrid(UIElement icon, string text)
        {
            RowDefinitions.Add(new RowDefinition());
            RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
            ColumnDefinitions.Add(new ColumnDefinition());

            this.textBlock = new TextBlock()
            {
                Text = text,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(4, 0, 4, 0)
            };
            Grid.SetRow(textBlock, 1);

            IsChecked = false;
            selection.Fill = new SolidColorBrush(Color.FromArgb(255, 150, 210, 255));
            selection.HorizontalAlignment = HorizontalAlignment.Stretch;
            selection.VerticalAlignment = VerticalAlignment.Stretch;
            Grid.SetRowSpan(selection, 2);

            Children.Add(selection);
            Children.Add(iconHolder);
            iconHolder.VerticalAlignment = VerticalAlignment.Center;
            iconHolder.HorizontalAlignment = HorizontalAlignment.Center;
            iconHolder.Margin = new Thickness(4, 4, 4, 0);

            this.Icon = icon;
            textBlock.FontSize = Settings.DefaultToolbarFontSize;
            Children.Add(textBlock);
        }

        bool isChecked;
        public bool IsChecked
        {
            get
            {
                return isChecked;
            }
            set
            {
                isChecked = value;
                selection.Opacity = value ? 1.0 : 0.01;
            }
        }

        Rectangle selection = new Rectangle();
        Grid iconHolder = new Grid();

        public double IconTextGap
        {
            set
            {
                iconHolder.Margin = new Thickness(4, 4, 4, value);
            }
        }

        public TextBlock textBlock;

        UIElement icon;
        public UIElement Icon
        {
            get
            {
                return icon;
            }
            set
            {
                if (icon != null)
                {
                    iconHolder.Children.Remove(icon);
                }
                icon = value;
                if (icon != null)
                {
                    iconHolder.Children.Add(icon);
                }
            }
        }

        public string Text
        {
            get
            {
                return textBlock.Text;
            }
            set
            {
                textBlock.Text = value;
            }
        }
    }
}