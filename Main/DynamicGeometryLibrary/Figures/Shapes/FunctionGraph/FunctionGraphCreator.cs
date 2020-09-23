using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.ComponentModel;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Coordinates)]
    [Order(1)]
    public class FunctionGraphCreator : Behavior
    {
        [PropertyGridName("Function graph")]
        public class Dialog
        {
            public Dialog(FunctionGraphCreator parent)
            {
                this.parent = parent;
            }

            FunctionGraphCreator parent;

            [PropertyGridVisible]
            [PropertyGridFocus]
            [PropertyGridEvent("KeyDown", "Func_KeyDown")]
            [PropertyGridName("f(x) = ")]
            public string Func { get; set; }

            internal void Func_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.Key == System.Windows.Input.Key.Escape)
                {
                    Cancel();
                    e.Handled = true;
                }
                else if (e.Key == System.Windows.Input.Key.Enter)
                {
                    Plot();
                    e.Handled = true;
                }
            }

            [PropertyGridVisible]
            public void Plot()
            {
                parent.PlotFunction(Func);
                Func = "";
            }

            [PropertyGridVisible]
            public void Cancel()
            {
                parent.AbortAndSetDefaultTool();
            }
        }

        public override object PropertyBag
        {
            get
            {
                if (PropertyDialog == null)
                {
                    PropertyDialog = new Dialog(this);
                }
                return PropertyDialog;
            }
        }

        protected Dialog PropertyDialog;

        protected void PlotFunction(string function)
        {
            var result = Compiler.Instance.CompileFunction(Drawing, function);
            Func<double, double> func = result.Function;
            if (func != null)
            {
                var graph = CreateFunctionGraph();
                graph.Drawing = Drawing;
                graph.FunctionText = function;
                Actions.Add(Drawing, graph);
                Drawing.ClearStatus();
            }
            else
            {
                Drawing.RaiseStatusNotification(result.ToString());
            }
        }

        protected virtual FunctionGraph CreateFunctionGraph()
        {
            return new FunctionGraph();
        }

        public override void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            PropertyDialog.Cancel();
        }

        public override string Name
        {
            get { return "Function"; }
        }

        public override string HintText
        {
            get
            {
                return "Enter an expression that depends on x, such as sin(x) or x * x - 3";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            var text = new TextBlock()
            {
                Text = "y=f(x)",
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
