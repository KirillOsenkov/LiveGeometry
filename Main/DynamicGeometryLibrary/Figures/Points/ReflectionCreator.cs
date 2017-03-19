using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;

namespace DynamicGeometry
{

#if !PLAYER
    [Category(BehaviorCategories.Transformations)]
    [Order(1)]
    public class ReflectionCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.Create<Point, Point>();   // Not used.
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
                if (Transformer.CanFigureBeMirrorForSource(result, FoundDependencies[0]))
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
            Check.ElementCount(FoundDependencies, 2);
            Check.NoNullElements(FoundDependencies);

            var results = Transformer.CreateReflectedFigure(
                Drawing,
                FoundDependencies[0],
                FoundDependencies[1]);

            Check.NotNull(results);
            Check.NoNullElements(results);
            foreach (IFigure f in results)
            {
                yield return f;
            }
        }

        public override string Name
        {
            get { return "Reflection"; }
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
                return "Select a mirror. The mirror can be a point, line, segment, ray, or circle (if the source is a point).";
            }
            return base.ConstructionHintText(args);
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Line(0.5, 0, 0.5, 1)
                .Polygon(
                    new SolidColorBrush(Colors.Yellow),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.4, 0.2),
                    new Point(0.4, 0.9),
                    new Point(0, 0.9))
                .Polygon(
                    new SolidColorBrush(Color.FromArgb(255, 128, 255, 128)),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.6, 0.2),
                    new Point(0.6, 0.9),
                    new Point(1, 0.9))
                .Canvas;
        }
    }
#endif
}