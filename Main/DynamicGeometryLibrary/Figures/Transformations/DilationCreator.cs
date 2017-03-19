using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Transformations)]
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
            [PropertyGridName("Factor = ")]
            public double factor { get; set; }

            [PropertyGridVisible]
            [PropertyGridName("Go")]
            public void Go()
            {
                if (parent.FoundDependencies.Count >= 2)
                {
                    parent.AddFiguresAndRestart();
                }
            }
        }

        DilationDialog dialog;

        public override object PropertyBag
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

        protected override IEnumerable<IFigure> CreateFigures()
        {
            Check.NotNull(FoundDependencies[0]);
            Check.NotNull(FoundDependencies[1]);

            var segment1 = (FoundDependencies.Count >= 3) ? FoundDependencies[2] : null;

            var results = Transformer.CreateDilatedFigure(
               Drawing,
               FoundDependencies[0],
               FoundDependencies[1],
               segment1,
               null,
               this.dialog.factor);

            Check.NotNull(results);
            Check.NoNullElements(results);
            foreach (IFigure f in results)
            {
                yield return f;
            }
        }

        public override string Name
        {
            get { return "Dilation"; }
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
                    new SolidColorBrush(Color.FromArgb(255, 128, 255, 128)),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.1, 0.9),
                    new Point(0.9, 0.9),
                    new Point(0.9, 0.1),
                    new Point(0.1, 0.1))
                .Polygon(
                    new SolidColorBrush(Colors.Yellow),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.1, 0.9),
                    new Point(0.5, 0.9),
                    new Point(0.5, 0.5),
                    new Point(0.1, 0.5))
                .Point(0.1, 0.9)
                .Canvas;
        }
    }
}