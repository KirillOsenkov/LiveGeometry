using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry
{
    /// <summary>
    /// Icon + caption. As a tool button: icon above the caption on a rounded plate that
    /// reacts to hover/press and shows the checked state. As a tab header: a small icon to
    /// the left of the caption, no plate (the tab strip draws its own selection).
    /// </summary>
    public class ButtonGrid : Grid
    {
        public ButtonGrid(UIElement icon, string text)
            : this(icon, text, false)
        {
        }

        public ButtonGrid(UIElement icon, string text, bool isTabHeader)
        {
            this.isTabHeader = isTabHeader;

            this.textBlock = new TextBlock()
            {
                Text = text,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = Settings.DefaultToolbarFontSize
            };

            iconHolder.VerticalAlignment = VerticalAlignment.Center;
            iconHolder.HorizontalAlignment = HorizontalAlignment.Center;

            if (isTabHeader)
            {
                ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
                ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
                textBlock.FontSize = Settings.DefaultToolbarFontSize + 1;
                textBlock.Margin = new Thickness(2, 0, 0, 0);
                Grid.SetColumn(textBlock, 1);

                // the full-size tool icon, shrunk
                var scaler = new LayoutTransformControl()
                {
                    LayoutTransform = new ScaleTransform(RibbonTheme.HeaderIconScale, RibbonTheme.HeaderIconScale),
                    Child = iconHolder,
                    VerticalAlignment = VerticalAlignment.Center
                };
                Children.Add(scaler);
                Children.Add(textBlock);
            }
            else
            {
                RowDefinitions.Add(new RowDefinition());
                RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
                ColumnDefinitions.Add(new ColumnDefinition());
                MinWidth = RibbonTheme.ButtonMinWidth;
                Margin = new Thickness(1, 0, 1, 0);

                plate.CornerRadius = RibbonTheme.ButtonCornerRadius;
                plate.BorderThickness = new Thickness(1);
                Grid.SetRowSpan(plate, 2);

                iconHolder.Margin = new Thickness(6, 5, 6, 0);
                textBlock.Margin = new Thickness(6, 1, 6, 4);
                textBlock.Foreground = RibbonTheme.Text;
                Grid.SetRow(textBlock, 1);

                Children.Add(plate);
                Children.Add(iconHolder);
                Children.Add(textBlock);

                PointerEntered += (s, e) => { isHovered = true; UpdatePlate(); };
                PointerExited += (s, e) => { isHovered = false; isPressed = false; UpdatePlate(); };
                PointerPressed += (s, e) => { isPressed = true; UpdatePlate(); };
                PointerReleased += (s, e) => { isPressed = false; UpdatePlate(); };
                UpdatePlate();
            }

            this.Icon = icon;
        }

        readonly bool isTabHeader;
        bool isHovered;
        bool isPressed;

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
                UpdatePlate();
            }
        }

        void UpdatePlate()
        {
            // never fully transparent-null: the plate is what makes the whole button clickable
            if (isChecked)
            {
                plate.Background = RibbonTheme.ButtonChecked;
                plate.BorderBrush = RibbonTheme.ButtonCheckedBorder;
            }
            else
            {
                plate.Background = isPressed
                    ? RibbonTheme.ButtonPressed
                    : isHovered ? RibbonTheme.ButtonHover : Brushes.Transparent;
                plate.BorderBrush = Brushes.Transparent;
            }
        }

        Border plate = new Border();
        Grid iconHolder = new Grid();

        public double IconTextGap
        {
            set
            {
                if (!isTabHeader)
                {
                    iconHolder.Margin = new Thickness(6, 5, 6, value);
                }
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
