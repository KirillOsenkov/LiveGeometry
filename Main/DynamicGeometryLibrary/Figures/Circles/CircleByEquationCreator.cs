using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Coordinates)]
    [Order(3)]
    public class CircleByEquationCreator : Behavior
    {
        [PropertyGridName("Circle equation")]
        public class Dialog
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
                if (e.Key == System.Windows.Input.Key.Escape)
                {
                    e.Handled = true;
                }
                else if (e.Key == System.Windows.Input.Key.Enter)
                {
                    AddCircle();
                    e.Handled = true;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Add circle")]
            public void AddCircle()
            {
                var xresult = parent.Drawing.CompileExpression(X);
                var yresult = parent.Drawing.CompileExpression(Y);
                var rresult = parent.Drawing.CompileExpression(R);

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