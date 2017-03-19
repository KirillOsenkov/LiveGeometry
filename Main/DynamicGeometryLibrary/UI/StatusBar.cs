using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

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

            border.Padding = new Thickness(4);
            border.Background = new SolidColorBrush(Color.FromArgb(255, 255, 255, 233));
            border.BorderBrush = new SolidColorBrush(Colors.Black);
            border.BorderThickness = new Thickness(1);
            border.MouseLeftButtonDown += border_MouseLeftButtonDown;

            TextBlock = new TextBlock();
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
                return Visibility == Visibility.Visible;
            }
            set
            {
                Visibility = value ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        void border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            IsVisible = false;
        }
    }
}
