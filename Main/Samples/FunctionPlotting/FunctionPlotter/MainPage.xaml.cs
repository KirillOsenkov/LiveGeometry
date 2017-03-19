using System;
using System.Windows.Controls;
using DynamicGeometry;

namespace FunctionPlotter
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
            Settings.Instance.ShowGrid = true;
            compiler = new Compiler();
            compiler.ExpressionTreeEvaluatorProvider = new ExpressionTreeCompiler();
            graph.ReadyForInteraction += graph_ReadyForInteraction;
        }

        void graph_ReadyForInteraction(object sender, EventArgs e)
        {
            functionText.Text = "x^2";
            plot = new FunctionGraph();
            graph.Drawing.Behavior = new Dragger();
            plot.Drawing = graph.Drawing;
            plot.FunctionText = functionText.Text;
            Actions.Add(graph.Drawing, plot);
        }

        private FunctionGraph plot;
        private Compiler compiler;

        private void functionText_TextChanged(object sender, TextChangedEventArgs e)
        {
            var result = compiler.CompileFunction(graph.Drawing, functionText.Text);
            Func<double, double> func = result.Function;

            plot.FunctionText = functionText.Text;

            if (func != null)
            {
                status.Text = "Ready";
            }
            else
            {
                status.Text = result.ToString();
            }
        }
    }
}
