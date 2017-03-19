using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Transformations)]
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
            [PropertyGridName("Angle = ")]
            public double angle { get; set; }

            [PropertyGridVisible]
            [PropertyGridName("Go")]
            public void Go()
            {
                if (parent.FoundDependencies.Count > 1)
                {
                    parent.AddFiguresAndRestart();
                }
            }
        }

        protected RotationDialog dialog;

        public override object PropertyBag
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

        protected override IEnumerable<IFigure> CreateFigures()
        {
            Check.NotNull(FoundDependencies[0]);
            Check.NotNull(FoundDependencies[1]);

            var angleProvider = (FoundDependencies.Count == 3) ? FoundDependencies[2] : null;

            var results = Transformer.CreateRotatedFigure(
                Drawing,
                FoundDependencies[0],
                FoundDependencies[1],
                angleProvider,
                this.dialog.angle);

            Check.NotNull(results);
            Check.NoNullElements(results);
            foreach (IFigure f in results)
            {
                yield return f;
            }
        }

        public override string Name
        {
            get { return "Rotation"; }
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
                    new SolidColorBrush(Colors.Yellow),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.3, 0.9),
                    new Point(0.9, 0.9),
                    new Point(0.9, 0.6))
                .Polygon(
                    new SolidColorBrush(Color.FromArgb(255, 128, 255, 128)),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.24, 0.06),
                    new Point(0.5, 0.21),
                    new Point(0.2, 0.73))
                .Canvas;
        }
    }
}