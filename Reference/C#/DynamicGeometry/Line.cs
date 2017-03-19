using System.Collections.Generic;
using System.Windows.Shapes;
using System.Windows.Controls;

namespace DynamicGeometry
{
    public abstract class LineBase : FigureBase
    {
        public LineBase()
        {
            Shape = Factory.CreateLineShape();
        }

        public Line Shape { get; set; }
        Canvas Parent;

        public PointPair Coordinates
        {
            get { return new PointPair { P1 = Point(0), P2 = Point(1) }; }
        }

        public override void OnPlacingOnContainer(System.Windows.Controls.Canvas newContainer)
        {
            Parent = newContainer;
            base.OnPlacingOnContainer(newContainer);
            newContainer.Children.Add(Shape);
        }

        public override void OnRemovingFromContainer(System.Windows.Controls.Canvas leavingContainer)
        {
            base.OnRemovingFromContainer(leavingContainer);
            leavingContainer.Children.Remove(Shape);
        }
    }

    public class LineTwoPoints : LineBase, ILine
    {
        public LineTwoPoints(IFigureList dependencies)
        {
            Dependencies = dependencies;
            RegisterWithDependencies();
        }

        public override ExpectedDependencyList RequiredDependencies
        {
            get { return ExpectedDependencyList.PointPoint; }
        }

        public override void Recalculate()
        {
            PointPair c = Coordinates;
            Shape.X1 = c.P1.X;
            Shape.Y1 = c.P1.Y;
            Shape.X2 = c.P2.X;
            Shape.Y2 = c.P2.Y;
        }

        public override IFigure HitTest(System.Windows.Point point)
        {
            return null;
        }
    }
}