using System;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Input;
using Avalonia.Media;
using System.ComponentModel;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Coordinates)]
    [Order(1)]
    public class FunctionGraphCreator : Behavior
    {
        [PropertyGridName("Function graph")]
        public class Dialog : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;

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
                if (e.Key == Avalonia.Input.Key.Escape)
                {
                    Cancel();
                    e.Handled = true;
                }
                else if (e.Key == Avalonia.Input.Key.Enter)
                {
                    Plot();
                    e.Handled = true;
                }
            }

            [PropertyGridVisible]
            [PropertyGridIcon(PropertyGridIcon.Check)]
            public void Plot()
            {
                // a function that doesn't compile stays in the box to be fixed (the status bar
                // says what's wrong); a plotted one leaves room for the next
                if (parent.PlotFunction(Func))
                {
                    Func = "";
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Func)));
                }
            }

            [PropertyGridVisible]
            [PropertyGridIcon(PropertyGridIcon.Cross)]
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

        /// <returns>Whether the function compiled and its graph was added</returns>
        protected bool PlotFunction(string function)
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
                return true;
            }

            Drawing.RaiseStatusNotification(result.ToString());
            return false;
        }

        protected virtual FunctionGraph CreateFunctionGraph()
        {
            return new FunctionGraph();
        }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
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
                FontSize = 13,
                Foreground = new SolidColorBrush(Colors.Black),
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
