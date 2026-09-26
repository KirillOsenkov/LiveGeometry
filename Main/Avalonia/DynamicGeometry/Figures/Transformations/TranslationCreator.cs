using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Media;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Transform)]
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

        /// <summary>A vector is the magnitude and the direction at once: nothing left to pick</summary>
        protected override Type GetExpectedDependencyType()
        {
            if (FoundDependencies.Count > 1 && FoundDependencies[1] is Vector)
            {
                return null;
            }

            return base.GetExpectedDependencyType();
        }

        protected override void AddFoundDependency(IFigure figure)
        {
            if (figure != null)
            {
                FoundDependencies.Add(figure);
            }
        }

        /// <summary>
        /// The magnitude is the second figure picked (a vector gives the direction too), the
        /// direction the third; what wasn't picked comes from the panel as a Number of its
        /// own, so that it can be edited later, and goes into the drawing before the points.
        /// </summary>
        protected override IEnumerable<IFigure> CreateFigures()
        {
            var source = FoundDependencies[0];
            Check.NotNull(source);
            IFigure magnitudeSource = null;
            IFigure directionSource = null;
            foreach (var picked in FoundDependencies.Skip(1))
            {
                if (picked is Vector)
                {
                    magnitudeSource = picked;
                    directionSource = picked;
                }
                else if (magnitudeSource == null)
                {
                    magnitudeSource = picked;
                }
                else if (directionSource == null)
                {
                    directionSource = picked;
                }
            }

            var typed = (TranslationDialog)PropertyBag;
            if (magnitudeSource == null)
            {
                magnitudeSource = Number.CreateAuxiliary(Drawing, typed.magnitude);
                yield return magnitudeSource;
            }

            if (directionSource == null)
            {
                directionSource = Number.CreateAuxiliary(Drawing, typed.direction);
                yield return directionSource;
            }

            var results = Transformer.CreateTranslatedFigure(
                Drawing,
                source,
                magnitudeSource,
                directionSource);

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