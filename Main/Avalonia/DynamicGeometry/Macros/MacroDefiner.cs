using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Input;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Misc)]
    [Order(10)]
    public class MacroDefiner : Behavior
    {
        public MacroDefiner()
        {

        }

        public override void Started()
        {
            dialog = new SelectInputsDialog(this);
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
            [PropertyGridIcon(PropertyGridIcon.Check)]
            public void OK()
            {
                // nothing to start from: nothing could be picked as a result either, and
                // Create tool would add a tool that makes nothing
                if (Parent.behavior.GetSelection().Count == 0)
                {
                    Parent.Drawing.RaiseStatusNotification("Click the figures the new tool starts from (two points, say), then press OK.");
                    return;
                }

                Parent.Inputs = Parent.behavior.GetSelection();
                Parent.behavior = new MacroResultSelector(Parent.Drawing, Parent.Inputs);
                Parent.Drawing.RaiseStatusNotification("Click the figures the new tool should make.");
                var dialog = new SelectResultsDialog(Parent);
                Parent.dialog = dialog;
                if (PropertyGrid != null)
                {
                    PropertyGrid.Show(dialog, null);
                }
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
            [PropertyGridName("Create tool")]
            [PropertyGridIcon(PropertyGridIcon.Check)]
            public void CreateTool()
            {
                if (Parent.behavior.GetSelection().Count == 0)
                {
                    Parent.Drawing.RaiseStatusNotification("Click the figures the new tool should make: those built on the ones picked before.");
                    return;
                }

                Parent.Results = Parent.behavior.GetSelection();
                var tool = Parent.CreateTool();

                // ready to use, with its name to change: Enter or OK keeps it
                var drawing = Parent.Drawing;
                drawing.Behavior = tool;
                drawing.RaiseDisplayProperties(tool.PropertyBag, focusProperty: nameof(UserDefinedTool.UserDefinedDialog.Name));
            }
        }

        public IList<IFigure> Inputs { get; set; }
        public IList<IFigure> Results { get; set; }

        /// <summary>
        /// The panel of the step at hand: the inputs, then the results. It is what the tool
        /// says its panel is, so that closing it (the cross) puts the tool down at either step.
        /// </summary>
        object dialog;

        FigureSelector behavior;

        public override object PropertyBag
        {
            get
            {
                if (dialog == null)
                {
                    dialog = new SelectInputsDialog(this);
                }

                return dialog;
            }
        }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            behavior?.Toggle(FindFigureToToggle(Coordinates(e)));
        }

        /// <summary>The figure a click here selects or lets go of: the first there is, or the one chosen with Tab</summary>
        IFigure FindFigureToToggle(Point coordinates)
        {
            return behavior == null ? null : Choice.Pick(behavior.FindFiguresToToggle(coordinates));
        }

        protected override IReadOnlyList<object> FindClickOptions(MouseEventArgs e)
        {
            return behavior == null ? System.Array.Empty<object>() : behavior.FindFiguresToToggle(Coordinates(e)).ToList<object>();
        }

        /// <summary>A halo on the figure a click would select or let go of, as tools show the figure they would take</summary>
        protected override IFigure GetFigureToPick(MouseEventArgs e)
        {
            return FindFigureToToggle(Coordinates(e));
        }

        protected override Cursor GetCursor(Point coordinates)
        {
            return FindFigureToToggle(coordinates) != null ? HandCursor : ArrowCursor;
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
                    nameof(AppTheme.ShapeIconFill),
                    nameof(AppTheme.ShapeOutline),
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

        // the order is the one the new tool will ask for them in
        public override string HintText
        {
            get { return "Click the inputs in order."; }
        }

        public virtual UserDefinedTool CreateTool()
        {
            // named after the first figure it makes (Catenary), a number added when that is taken
            var firstName = Results.Select(r => r.Name).FirstOrDefault(n => !n.IsEmpty()) ?? "Custom tool";
            foreach (var result in Results.ToArray())
            {
                AddIntermediateResults(result);
            }
            Results = Sort(Results);
            string macro = MacroSerializer.WriteMacroToString(Inputs, Results, UniqueToolName(firstName));
            return UserDefinedTool.AddFromString(macro);
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