using System.Windows;

namespace DynamicGeometry
{
    public abstract class ControlBase : CoordinatesShapeBase<FrameworkElement>
    {
        protected override int DefaultZOrder()
        {
            return (int)ZOrder.Controls;
        }

        public override IFigure HitTest(Point point)
        {
            if (Rect.Contains(point))
            {
                return this;
            }

            return null;
        }

        public override void ApplyStyle()
        {
            if (this.Style == null)
            {
                return;
            }

            if (Drawing != null)
            {
                UpdateVisual();
            }
        }

        public override void UpdateVisual()
        {
            shape.MoveTo(ToPhysical(Coordinates));
        }

        public System.Windows.Rect Rect
        {
            get
            {
                var rect = new System.Windows.Rect();
                rect.X = Coordinates.X;

                // Factor in possible scale transform.
                var p = Shape.RenderTransform.Transform(new Point(Shape.ActualWidth, Shape.ActualHeight));
 
                rect.Width = ToLogical(p.X);
                rect.Height = ToLogical(p.Y);
                rect.Y = Coordinates.Y - rect.Height;
                return rect;
            }
        }
    }
}
