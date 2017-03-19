using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Coordinates)]
    [Order(2)]
    public class LineByEquationCreator : Behavior
    {
        [PropertyGridName("y = mx + b")]
        public class Dialog
        {
            public Dialog(LineByEquationCreator parent)
            {
                this.parent = parent;
            }

            LineByEquationCreator parent;

            [PropertyGridVisible]
            [PropertyGridFocus]
            [PropertyGridEvent("KeyDown", "Common_KeyDown")]
            [PropertyGridName("m = ")]
            public string m { get; set; }

            [PropertyGridVisible]
            [PropertyGridEvent("KeyDown", "Common_KeyDown")]
            [PropertyGridName("b = ")]
            public string b { get; set; }

            internal void Common_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.Key == System.Windows.Input.Key.Escape)
                {
                    e.Handled = true;
                }
                else if (e.Key == System.Windows.Input.Key.Enter)
                {
                    AddLine();
                    e.Handled = true;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Add line")]
            public void AddLine()
            {
                parent.AddLine(m, b);
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

        public virtual void AddLine(string m, string b)
        {
            var mresult = Drawing.CompileExpression(m);
            var bresult = Drawing.CompileExpression(b);

            if (mresult.IsSuccess && bresult.IsSuccess)
            {
                var line = Factory.CreateLineByEquation(Drawing, mresult.Dependencies.Union(bresult.Dependencies).ToList());
                var equation = new SlopeInterseptLineEquation(line, m, b);
                line.Equation = equation;
                equation.Recalculate();
                Actions.Add(Drawing, line);
            }
        }

        public override string Name
        {
            get { return "Line"; }
        }

        public override string HintText
        {
            get { return "Enter expressions for slope (m) and y-intercept (b) of the line."; }
        }

        public override FrameworkElement CreateIcon()
        {
            var text = new TextBlock()
            {
                Text = "y=mx+b",
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