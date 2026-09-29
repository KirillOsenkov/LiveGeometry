using System.ComponentModel;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Input;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Coordinates)]
    [Order(3)]
    public class CircleByEquationCreator : Behavior
    {
        [PropertyGridName("Circle equation")]
        public class Dialog : ToolPanel
        {
            public Dialog(CircleByEquationCreator parent)
            {
                this.parent = parent;
            }

            CircleByEquationCreator parent;

            [PropertyGridFocus]
            [PropertyGridVisible]
            [PropertyGridEvent("KeyDown", "Common_KeyDown")]
            [PropertyGridName("Center X = ")]
            public string X { get; set; }

            [PropertyGridVisible]
            [PropertyGridEvent("KeyDown", "Common_KeyDown")]
            [PropertyGridName("Center Y = ")]
            public string Y { get; set; }

            [PropertyGridVisible]
            [PropertyGridEvent("KeyDown", "Common_KeyDown")]
            [PropertyGridName("Radius = ")]
            public string R { get; set; }

            internal void Common_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.Key == Avalonia.Input.Key.Escape)
                {
                    e.Handled = true;
                }
                else if (e.Key == Avalonia.Input.Key.Enter)
                {
                    AddCircle();
                    e.Handled = true;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Add circle")]
            [PropertyGridIcon(PropertyGridIcon.Plus)]
            public void AddCircle()
            {
                var xresult = Compile(parent.Drawing, nameof(X), X);
                var yresult = Compile(parent.Drawing, nameof(Y), Y);
                var rresult = Compile(parent.Drawing, nameof(R), R);

                if (xresult.IsSuccess && yresult.IsSuccess && rresult.IsSuccess)
                {
                    var circle = parent.CreateCircle(X, Y, R);
                    Actions.Add(parent.Drawing, circle);
                }
            }
        }

        Dialog dialog;

        public override object PropertyBag
        {
            get
            {
                if (dialog == null)
                {
                    dialog = new Dialog(this);
                }
                return dialog;
            }
        }

        public virtual IFigure CreateCircle(string X, string Y, string R)
        {
            return Factory.CreateCircleByEquation(Drawing, X, Y, R);
        }

        public override string Name
        {
            get
            {
                return "Circle";
            }
        }

        public override string HintText
        {
            get
            {
                return "Enter expressions for the center and radius of the circle.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            var text = new TextBlock()
            {
                Text = "x²+y²=r²",
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