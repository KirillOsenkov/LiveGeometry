using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Coordinates)]
    [Order(2)]
    public class LineByEquationCreator : Behavior
    {
        [PropertyGridName("y = mx + b")]
        public class Dialog : ToolPanel
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
                if (e.Key == Avalonia.Input.Key.Escape)
                {
                    e.Handled = true;
                }
                else if (e.Key == Avalonia.Input.Key.Enter)
                {
                    AddLine();
                    e.Handled = true;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Add line")]
            [PropertyGridIcon(PropertyGridIcon.Plus)]
            public void AddLine()
            {
                var mresult = Compile(parent.Drawing, nameof(m), m);
                var bresult = Compile(parent.Drawing, nameof(b), b);
                if (mresult.IsSuccess && bresult.IsSuccess)
                {
                    parent.AddLine(m, b, mresult.Dependencies.Union(bresult.Dependencies).ToList());
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

        /// <summary>Adds the line of a slope and an intercept that compile</summary>
        public virtual void AddLine(string m, string b, IList<IFigure> dependencies)
        {
            var line = Factory.CreateLineByEquation(Drawing, dependencies);
            var equation = new SlopeInterseptLineEquation(line, m, b);
            line.Equation = equation;
            equation.Recalculate();
            Actions.Add(Drawing, line);
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