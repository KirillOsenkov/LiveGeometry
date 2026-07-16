using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;

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

        public Avalonia.Rect Rect
        {
            get
            {
                var rect = new Avalonia.Rect();
                rect = rect.WithX(Coordinates.X);

                // Factor in possible scale transform.
                var transform = Shape.RenderTransform?.Value ?? Avalonia.Matrix.Identity;
                var p = new Point(Shape.ActualWidth, Shape.ActualHeight).Transform(transform);
 
                rect = rect.WithWidth(ToLogical(p.X));
                rect = rect.WithHeight(ToLogical(p.Y));
                rect = rect.WithY(Coordinates.Y - rect.Height);
                return rect;
            }
        }
    }
}
