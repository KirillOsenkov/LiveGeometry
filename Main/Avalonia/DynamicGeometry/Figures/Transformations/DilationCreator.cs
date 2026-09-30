using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;
using Avalonia.Input;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Transform)]
    [Order(4)]
    public class DilationCreator : FigureCreator
    {
        [PropertyGridName("Dilation Factor")]
        public class DilationDialog
        {
            public DilationDialog(DilationCreator parent)
            {
                this.parent = parent;
            }

            DilationCreator parent;

            [PropertyGridVisible]
            [PropertyGridFocus]
            [PropertyGridEvent("KeyDown", "Factor_KeyDown")]
            [PropertyGridName("Factor = ")]
            public double factor { get; set; } = 2; // 0 would squash the figure into the center

            /// <summary>Enter is OK, unless what was typed is no number (the box says so, and the factor is still the old one)</summary>
            public void Factor_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.Key == Key.Enter)
                {
                    if (!(sender is StringEditor editor) || string.IsNullOrEmpty(editor.ErrorText))
                    {
                        OK();
                    }

                    e.Handled = true;
                }
            }

            [PropertyGridVisible]
            [PropertyGridIcon(PropertyGridIcon.Check)]
            public void OK()
            {
                if (parent.FoundDependencies.Count >= 2)
                {
                    parent.AddFiguresAndRestart();
                }
            }
        }

        DilationDialog dialog;

        /// <summary>The typed factor, kept from one dilation to the next</summary>
        DilationDialog Dialog
        {
            get
            {
                if (dialog == null)
                {
                    dialog = new DilationDialog(this);
                }
                return dialog;
            }
        }

        /// <summary>
        /// Only at the step that asks for the factor: shown from the start, it took the keyboard
        /// while the source and the center were still to be clicked. "Construction complete" is
        /// raised before the found figures are cleared, but after the transaction is gone.
        /// </summary>
        public override object PropertyBag
        {
            get
            {
                return Transaction != null && FoundDependencies.Count == 2 ? Dialog : null;
            }
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.Create<IFigure, IPoint, ILengthProvider>();
        }

        protected override IFigure LookForExpectedDependencyUnderCursor(Point coordinates)
        {
            if (FoundDependencies.Count == 0)
            {
                var result = Drawing.Figures.HitTest(coordinates);
                if (Transformer.CanBeTransformSource(result))
                {
                    return result;
                }
            }
            else if (FoundDependencies.Count == 1)
            {
                var result = Drawing.Figures.HitTest(coordinates);
                if (result is IPoint)
                {
                    return result;
                }
            }
            else if (FoundDependencies.Count == 2)
            {
                var result = Drawing.Figures.HitTest(coordinates);
                if (result is ILengthProvider)
                {
                    return result;
                }
            }
            return base.LookForExpectedDependencyUnderCursor(coordinates);
        }

        protected override bool ExpectingAPoint()
        {
            return false;
        }

        protected override void AddFoundDependency(IFigure figure)
        {
            if (figure != null)
            {
                FoundDependencies.Add(figure);
            }
        }

        /// <summary>
        /// A typed factor becomes a Number of its own, before the points and shared by every
        /// point of the dilated figure; a clicked figure with a length is the source as it is.
        /// </summary>
        protected override IEnumerable<IFigure> CreateFigures()
        {
            Check.NotNull(FoundDependencies[0]);
            Check.NotNull(FoundDependencies[1]);

            var factor = FoundDependencies.Count >= 3 ? FoundDependencies[2] : null;
            if (factor == null)
            {
                factor = Number.CreateAuxiliary(Drawing, Dialog.factor);
                yield return factor;
            }

            var results = Transformer.CreateDilatedFigure(
               Drawing,
               FoundDependencies[0],
               FoundDependencies[1],
               factor,
               lengthProvider2: null);

            Check.NotNull(results);
            Check.NoNullElements(results);
            foreach (IFigure f in results)
            {
                yield return f;
            }
        }

        /// <summary>A factor set in the panel after a dilation is the next dilation's</summary>
        protected override void TakeDefaultsFrom(ITiedValues created)
        {
            if (created is DilatedPoint point && point.FactorSource is Number number)
            {
                Dialog.factor = number.Value;
            }
        }

        protected override string CreatedFigureHint(ITiedValues values)
        {
            return values.IsTied(nameof(DilatedPoint.Factor))
                ? "click another segment or a slider to take the factor from it instead."
                : "set its factor in the panel, or click a segment or a slider to take the factor from it.";
        }

        public override string Name
        {
            get { return "Dilate"; }
        }

        public override string HintText
        {
            get
            {
                return "Select the source figure.";
            }
        }

        public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
        {
            if (FoundDependencies.Count == 0)
            {
                return "Select the source figure.";
            }
            else if (FoundDependencies.Count == 1)
            {
                return "Select a point to use for the center of dilation.";
            }
            else if (FoundDependencies.Count == 2)
            {
                return "Define the dilation factor by selecting a figure with length or entering a value.";
            }
            return base.ConstructionHintText(args);
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Polygon(
                    nameof(AppTheme.ImageFill),
                    nameof(AppTheme.Ink),
                    new Point(0.1, 0.9),
                    new Point(0.9, 0.9),
                    new Point(0.9, 0.1),
                    new Point(0.1, 0.1))
                .Polygon(
                    nameof(AppTheme.ShapeIconFill),
                    nameof(AppTheme.Ink),
                    new Point(0.1, 0.9),
                    new Point(0.5, 0.9),
                    new Point(0.5, 0.5),
                    new Point(0.1, 0.5))
                .Point(0.1, 0.9)
                .Canvas;
        }
    }
}