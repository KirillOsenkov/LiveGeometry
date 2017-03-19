using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Misc)]
    [Order(3)]
    public class LabelCreator : Behavior
    {
        public override void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var label = Factory.CreateLabel(Drawing);
            label.Text = "Text";
            label.MoveTo(Coordinates(e));
            Actions.Add(Drawing, label);
            var drawing = Drawing;
            AbortAndSetDefaultTool();
            drawing.RaiseStatusNotification("");
            drawing.RaiseDisplayProperties(label);
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
                Text = "Abc",
                FontStyle = FontStyles.Italic,
                FontWeight = FontWeights.Bold,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
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