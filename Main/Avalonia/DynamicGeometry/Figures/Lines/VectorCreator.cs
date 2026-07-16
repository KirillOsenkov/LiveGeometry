using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;
using Avalonia.Media;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Coordinates)]
    [Order(4)]
    public class VectorCreator : FigureCreator
    {
        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateVector(Drawing, FoundDependencies);
        }

        public override string Name
        {
            get { return "Vector"; }
        }

        public override string HintText
        {
            get
            {
                return "Click (and release) twice to connect two points with a vector.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Line(0.25, 0.75, 0.75, 0.25)
                .Polygon(new SolidColorBrush(Colors.Black), new SolidColorBrush(Colors.Black), new Point(0.75, 0.25), new Point(0.5, 0.4), new Point(0.6, 0.5))
                .Canvas;
        }
    }
}