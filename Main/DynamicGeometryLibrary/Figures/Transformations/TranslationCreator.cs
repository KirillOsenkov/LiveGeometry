using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using System.Linq;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Transformations)]
    [Order(3)]
    public class TranslationCreator : FigureCreator
    {
        [PropertyGridName("Translation Values")]
        public class TranslationDialog
        {
            public TranslationDialog(TranslationCreator parent)
            {
                this.parent = parent;
            }

            TranslationCreator parent;

            [PropertyGridVisible]
            [PropertyGridFocus]
            [PropertyGridName("Magnitude = ")]
            public double magnitude { get; set; }

            [PropertyGridVisible]
            [PropertyGridName("Direction = ")]
            public double direction { get; set; }

            [PropertyGridVisible]
            [PropertyGridName("Go")]
            public void Go()
            {
                if (parent.FoundDependencies.Count > 0)
                {
                    parent.AddFiguresAndRestart();
                }
            }
        }

        TranslationDialog dialog;

        public override object PropertyBag
        {
            get
            {
                if (dialog == null)
                {
                    dialog = new TranslationDialog(this);
                }
                return dialog;
            }
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.Create<Point, Point, Point>();    // Using number of dependencies only.
        }

        protected override IFigure LookForExpectedDependencyUnderCursor(Point coordinates)
        {
            var result = Drawing.Figures.HitTest(coordinates);
            if (FoundDependencies.Count == 0)
            {
                if (Transformer.CanBeTransformSource(result))
                {
                    return result;
                }
            }
            else if (FoundDependencies.Count == 1)
            {
                if (result is Vector || result is ILengthProvider)
                {
                    return result;
                }
            }
            else if (FoundDependencies.Count == 2)
            {
                if (result is Vector || result is IAngleProvider)
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
            var source = FoundDependencies[0];
            Check.NotNull(source);
            var dependenciesSubset = new List<IFigure>(FoundDependencies.Without(source));
            var results = Transformer.CreateTranslatedFigure(
                Drawing,
                source,
                dependenciesSubset,
                dialog.magnitude,
                dialog.direction);

            Check.NotNull(results);
            Check.NoNullElements(results);
            foreach (IFigure f in results)
            {
                yield return f;
            }
        }

        public override string Name
        {
            get { return "Translation"; }
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
                return "Define the translation magnitude by selecting a vector, figure with length, or entering the value.";
            }
            else if (FoundDependencies.Count == 2)
            {
                return "Define the translation direction by selecting a vector, figure with an angle or entering the value.";
            }
            return base.ConstructionHintText(args);
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Polygon(
                    new SolidColorBrush(Colors.Yellow),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.1, 0.9),
                    new Point(0.4, 0.9),
                    new Point(0.4, 0.6),
                    new Point(0.1, 0.6))
                .Polygon(
                    new SolidColorBrush(Color.FromArgb(255, 128, 255, 128)),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.6, 0.4),
                    new Point(0.9, 0.4),
                    new Point(0.9, 0.1),
                    new Point(0.6, 0.1))
                .Canvas;
        }
    }
}