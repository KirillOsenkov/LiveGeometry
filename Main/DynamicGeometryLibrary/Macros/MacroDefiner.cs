using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Custom)]
    [Order(1)]
    public class MacroDefiner : Behavior
    {
        public MacroDefiner()
        {

        }

        public override void Started()
        {
            selectInputsDialog = new SelectInputsDialog(this);
            behavior = new MacroInputSelector() { Drawing = Drawing };
        }

        public override void Stopping()
        {
            Drawing.Figures.EnableAll();
            Drawing.Figures.ClearSelection();
            base.Stopping();
        }

        [PropertyGridName("Select input figures")]
        public class SelectInputsDialog : IPropertyGridHost
        {
            public SelectInputsDialog(MacroDefiner parent)
            {
                Parent = parent;
            }

            MacroDefiner Parent;

            [PropertyGridVisible]
            public void Done()
            {
                Parent.Inputs = Parent.behavior.GetSelection();
                Parent.behavior = new MacroResultSelector(Parent.Drawing, Parent.Inputs);
                var dialog = new SelectResultsDialog(Parent);
                if (PropertyGrid != null)
                {
                    PropertyGrid.Show(dialog, null);
                }
            }

            [PropertyGridVisible]
            public void Cancel()
            {
                Parent.AbortAndSetDefaultTool();
            }

            public PropertyGrid PropertyGrid { get; set; }
        }

        [PropertyGridName("Now select resulting figures")]
        public class SelectResultsDialog
        {
            public SelectResultsDialog(MacroDefiner parent)
            {
                Parent = parent;
            }

            MacroDefiner Parent;

            [PropertyGridVisible]
            [PropertyGridName("Done - create a tool")]
            public void Done()
            {
                Parent.Results = Parent.behavior.GetSelection();
                Parent.CreateTool();
                Parent.AbortAndSetDefaultTool();
            }

            [PropertyGridVisible]
            public void Cancel()
            {
                Parent.AbortAndSetDefaultTool();
            }
        }

        public IList<IFigure> Inputs { get; set; }
        public IList<IFigure> Results { get; set; }

        SelectInputsDialog selectInputsDialog;

        FigureSelector behavior;

        public override object PropertyBag
        {
            get
            {
                if (selectInputsDialog == null)
                {
                    selectInputsDialog = new SelectInputsDialog(this);
                }
                return selectInputsDialog;
            }
        }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (behavior != null)
            {
                behavior.MouseDown(sender, e);
            }
        }

        public override void KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Restart();
                e.Handled = true;
            }
        }

        public override FrameworkElement CreateIcon()
        {
            double a = 0.2, b = 0.4, c = 0.6, d = 0.8;
            return IconBuilder.BuildIcon()
                .Polygon(
                    Factory.CreateLinearGradient(
                        Color.FromArgb(255, 200, 255, 0),
                        Color.FromArgb(255, 255, 255, 0),
                        90),
                    new SolidColorBrush(Colors.Black),
                    new Point(a, b),
                    new Point(b, b),
                    new Point(b, a),
                    new Point(c, a),
                    new Point(c, b),
                    new Point(d, b),
                    new Point(d, c),
                    new Point(c, c),
                    new Point(c, d),
                    new Point(b, d),
                    new Point(b, c),
                    new Point(a, c))
                .Canvas;
        }

        public override string Name
        {
            get { return "Define figure"; }
        }

        public virtual void CreateTool()
        {
            foreach (var result in Results.ToArray())
            {
                AddIntermediateResults(result);
            }
            Results = Sort(Results);
            string macro = MacroSerializer.WriteMacroToString(Inputs, Results);
            UserDefinedTool.AddFromString(macro);
        }

        protected IList<IFigure> Sort(IList<IFigure> set)
        {
            IList<IFigure> result = new List<IFigure>();

            var sorted = set.TopologicalSort(f => f.Dependencies);
            foreach (var item in sorted)
            {
                if (set.Contains(item))
                {
                    result.Add(item);
                }
            }

            return result;
        }

        protected void AddIntermediateResults(IFigure figure)
        {
            if (Inputs.Contains(figure))
            {
                return;
            }
            if (!Results.Contains(figure))
            {
                Results.Insert(0, figure);
            }
            foreach (var dependency in figure.Dependencies)
            {
                AddIntermediateResults(dependency);
            }
        }
    }
}