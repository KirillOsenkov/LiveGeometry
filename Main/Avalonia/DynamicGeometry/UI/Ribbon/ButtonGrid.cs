using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry
{
    /// <summary>
    /// Icon + caption. As a tool button: icon above the caption on a rounded plate that
    /// reacts to hover/press and shows the checked state. As a group (tab) header: the same
    /// layout, and instead of the plate a <see cref="TabOutline"/> that draws the tab shape.
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
                RowDefinitions.Add(new RowDefinition());
                RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
                ColumnDefinitions.Add(new ColumnDefinition());
                MinWidth = RibbonTheme.ButtonMinWidth;

                // The tab's feet flare outwards, and drawing outside of one's bounds gets clipped,
                // so the header is wider than the tab body by the flare on both sides.
                Grid.SetRowSpan(tabOutline, 2);

                double side = TabOutline.Flare + 5;
                iconHolder.Margin = new Thickness(side, 6, side, 0);
                textBlock.Margin = new Thickness(side, 1, side, 5);
                textBlock.FontWeight = FontWeight.Medium;
                Grid.SetRow(textBlock, 1);

                Children.Add(tabOutline);
                Children.Add(iconHolder);
                Children.Add(textBlock);

                // the grid itself must be hit-testable over its whole area
                Background = Brushes.Transparent;
                PointerEntered += (s, e) => tabOutline.IsHovered = true;
                PointerExited += (s, e) => tabOutline.IsHovered = false;
                UpdatePlate();
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
            if (isTabHeader)
            {
                textBlock.Foreground = isChecked ? RibbonTheme.TabHeaderTextSelected : RibbonTheme.TabHeaderText;
                tabOutline.IsSelected = isChecked;
                return;
            }

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
        TabOutline tabOutline = new TabOutline();
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
