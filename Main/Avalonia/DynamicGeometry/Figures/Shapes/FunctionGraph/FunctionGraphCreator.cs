using System.ComponentModel;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Coordinates)]
    [Order(1)]
    public class FunctionGraphCreator : Behavior
    {
        [PropertyGridName("Function graph")]
        public class Dialog : ToolPanel, INotifyPropertyChanged
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
                // a function that doesn't compile stays in the box to be fixed, with what's
                // wrong under it; a plotted one leaves room for the next
                var result = Compiler.Instance.CompileFunction(parent.Drawing, Func);
                ReportError(nameof(Func), result.GetErrorText(whenEmpty: EmptyFunctionError));
                if (result.IsSuccess)
                {
                    parent.PlotFunction(Func);
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

        public const string EmptyFunctionError = "Type an expression in x, such as sin(x).";

        /// <summary>Adds the graph of a function that compiles</summary>
        protected void PlotFunction(string function)
        {
            var graph = CreateFunctionGraph();
            graph.Drawing = Drawing;
            graph.FunctionText = function;
            Actions.Add(Drawing, graph);
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
