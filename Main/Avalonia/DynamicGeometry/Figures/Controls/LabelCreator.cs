using System.ComponentModel;
using Avalonia.Controls;
using Avalonia.Layout;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Misc)]
    [Order(3)]
    public class LabelCreator : Behavior
    {
        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            var label = Factory.CreateLabel(Drawing);
            label.Text = "Text";
            label.MoveTo(Coordinates(e));
            Actions.Add(Drawing, label);
            var drawing = Drawing;
            AbortAndSetDefaultTool();
            drawing.RaiseStatusNotification("");
            drawing.RaiseDisplayProperties(label, focusProperty: nameof(label.Text));
        }

        public override string Name
        {
            get { return "Text"; }
        }

        public override string HintText
        {
            get
            {
                return "Click to add a text label.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            var text = new TextBlock()
            {
                // as the icons of the Coordinates tab that are text (y=f(x)): same size, upright, in ink
                Text = "Abc",
                FontSize = 13,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            text.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Ink));
            var grid = new Grid()
            {
                MinWidth = IconBuilder.IconSize,
                MinHeight = IconBuilder.IconSize,
            };
            grid.Children.Add(text);
            return grid;
        }
    }
}