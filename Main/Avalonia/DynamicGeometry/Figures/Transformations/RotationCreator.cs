using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;
using Avalonia.Input;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Transform)]
    [Order(2)]
    public class RotationCreator : FigureCreator
    {
        [PropertyGridName("Rotation Angle")]
        public partial class RotationDialog
        {
            public RotationDialog(RotationCreator parent)
            {
                this.parent = parent;
            }

            RotationCreator parent;

            [PropertyGridVisible]
            [PropertyGridFocus]
            [PropertyGridEvent("KeyDown", "Angle_KeyDown")]
            [PropertyGridName("Angle = ")]
            public double angle { get; set; }

            /// <summary>Enter is OK, unless what was typed is no number (the box says so, and the angle is still the old one)</summary>
            public void Angle_KeyDown(object sender, KeyEventArgs e)
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
                if (parent.FoundDependencies.Count > 1)
                {
                    parent.AddFiguresAndRestart();
                }
            }
        }

        protected RotationDialog dialog;

        /// <summary>The typed angle, kept from one rotation to the next</summary>
        RotationDialog Dialog
        {
            get
            {
                if (dialog == null)
                {
                    dialog = new RotationDialog(this);
                }
                return dialog;
            }
        }

        /// <summary>
        /// Only at the step that asks for the angle: shown from the start, it took the keyboard
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
            return DependencyList.Create<IFigure, IPoint, IAngleProvider>();
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
                if (result is IAngleProvider)
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
        /// A typed angle becomes a Number of its own, before the points and shared by every
        /// point of the rotated figure; a clicked figure with an angle is the source as it is.
        /// </summary>
        protected override IEnumerable<IFigure> CreateFigures()
        {
            Check.NotNull(FoundDependencies[0]);
            Check.NotNull(FoundDependencies[1]);

            var angle = FoundDependencies.Count == 3 ? FoundDependencies[2] : null;
            if (angle == null)
            {
                angle = Number.CreateAuxiliary(Drawing, Dialog.angle);
                yield return angle;
            }

            var results = Transformer.CreateRotatedFigure(
                Drawing,
                FoundDependencies[0],
                FoundDependencies[1],
                angle);

            Check.NotNull(results);
            Check.NoNullElements(results);
            foreach (IFigure f in results)
            {
                yield return f;
            }
        }

        /// <summary>An angle set in the panel after a rotation is the next rotation's</summary>
        protected override void TakeDefaultsFrom(ITiedValues created)
        {
            if (created is RotatedPoint point && point.AngleSource is Number number)
            {
                Dialog.angle = number.Value;
            }
        }

        protected override string CreatedFigureHint(ITiedValues values)
        {
            return values.IsTied(nameof(RotatedPoint.Angle))
                ? "click another angle or a slider to take the angle from it instead."
                : "set its angle in the panel, or click an angle or a slider to take the angle from it.";
        }

        public override string Name
        {
            get { return "Rotate"; }
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
                return "Select a point to use for the center of rotation.";
            }
            else if (FoundDependencies.Count == 2)
            {
                return "Define the angle by selecting a figure with an angle (such as an arc or angle measurement) or entering the value.";
            }
            return base.ConstructionHintText(args);
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Point(0.1, 0.9)
                .Polygon(
                    nameof(AppTheme.ShapeIconFill),
                    nameof(AppTheme.Ink),
                    new Point(0.3, 0.9),
                    new Point(0.9, 0.9),
                    new Point(0.9, 0.6))
                .Polygon(
                    nameof(AppTheme.ImageFill),
                    nameof(AppTheme.Ink),
                    new Point(0.24, 0.06),
                    new Point(0.5, 0.21),
                    new Point(0.2, 0.73))
                .Canvas;
        }
    }
}