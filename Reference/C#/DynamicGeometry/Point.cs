using System.Windows;
using System.Collections.Generic;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public abstract class PointBase : FigureBase, IPoint
    {
        public PointBase()
        {
            Shape = Factory.CreatePointShape();
        }

        public virtual void MoveToCore(Point newLocation)
        {
            Coordinates = newLocation;
            Shape.MoveTo(Coordinates);
        }

        public override void OnPlacingOnContainer(System.Windows.Controls.Canvas newContainer)
        {
            Parent = newContainer;
            base.OnPlacingOnContainer(newContainer);
            newContainer.Children.Add(Shape);
            Shape.SetValue(Canvas.ZIndexProperty, 1);
        }

        public override void OnRemovingFromContainer(System.Windows.Controls.Canvas leavingContainer)
        {
            base.OnRemovingFromContainer(leavingContainer);
            leavingContainer.Children.Remove(Shape);
        }

        protected Point Coordinates;

        Point IPoint.Coordinates
        {
            get { return Coordinates; }
        }

        public Shape Shape { get; set; }
        Canvas Parent;

        public override IFigure HitTest(Point point)
        {
            var result = VisualTreeHelper.HitTest(Parent, point);
            FrameworkElement ell = result.VisualHit as FrameworkElement;
            if (ell == Shape)
            {
                return this;
            }
            return null;
        }
    }

    public class FreePoint : PointBase
    {
        public FreePoint()
        {

        }

        public FreePoint(Point initialPosition)
        {
            MoveTo(initialPosition);
        }

        public void MoveTo(Point newPosition)
        {
            MoveToCore(newPosition);
        }

        public override ExpectedDependencyList RequiredDependencies
        {
            get { return ExpectedDependencyList.None; }
        }

        public override void Recalculate()
        {
            //MoveTo(Coordinates + new Vector(1, 2));
        }
    }

    public class MidPoint : PointBase, IPoint
    {
        public MidPoint(IFigureList dependencies)
        {
            Dependencies = dependencies;
            RegisterWithDependencies();
        }

        public override void Recalculate()
        {
            Coordinates.X = (Point(0).X + Point(1).X) / 2;
            Coordinates.Y = (Point(0).Y + Point(1).Y) / 2;
            MoveToCore(Coordinates);
        }

        public override ExpectedDependencyList RequiredDependencies
        {
            get { return ExpectedDependencyList.PointPoint; }
        }
    }
}

