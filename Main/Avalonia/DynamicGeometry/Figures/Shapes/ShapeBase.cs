using Avalonia;
using Avalonia.Controls;

namespace DynamicGeometry
{
    public partial interface IShapeWithInterior
    {
        double Area { get; }
    }

    public abstract partial class ShapeBase<TShape> : FigureBase
        where TShape : FrameworkElement
    {
        public ShapeBase()
        {
            Shape = CreateShape();
            ZIndex = DefaultZOrder();
        }

        protected virtual int DefaultZOrder()
        {
            return (int)ZOrder.Figures;
        }

        protected TShape shape;
        public TShape Shape
        {
            get
            {
                return shape;
            }
            set
            {
                shape = value;
                if (shape != null)
                {
                    shape.ZIndex = ZIndex;
                }
            }
        }

        public override int ZIndex
        {
            get
            {
                return base.ZIndex;
            }
            set
            {
                base.ZIndex = value;
                if (shape != null)
                {
                    shape.ZIndex = ZIndex;
                }
            }
        }

        protected abstract TShape CreateShape();

        /// <summary>
        /// Implementation of the IMovable.MoveTo method
        /// </summary>
        /// <example>
        /// Usually just MoveToCore (set coordinates) and UpdateVisual
        /// </example>
        /// <param name="newPosition">Coordinates to move this shape to</param>
        public void MoveTo(Point newPosition)
        {
            MoveToCore(newPosition);
            UpdateVisual();
        }

        public virtual bool AllowMove()
        {
            return !this.Locked && this.Dependencies.IsEmpty();
        }

        public virtual void MoveToCore(Point newLocation)
        {
        }

        public override void UpdateVisual()
        {
        }

        public override bool Exists
        {
            get
            {
                return mExists;
            }
            set
            {
                if (mExists == value)
                {
                    return;
                }

                mExists = value;
                UpdateShapeVisibility();
            }
        }

        public IFigure HitTestShape(Avalonia.Point point)
        {
            // Prevent error caused by calling Shape.TransformToVisual() with the Shape not yet rendered.  Occurs when figure is not in Drawing.
            // An example of when this occurs is when a Drawing containing a PointOnFigure on an Arc is loaded.
            if (Avalonia.VisualTree.VisualExtensions.GetVisualParent(Shape) == null)
            {
                return null;
            }

            // a hidden figure shown while selected (IsGhost) is still hidden
            if (!Visible)
            {
                return null;
            }

            var oldHitTestVisible = Shape.IsHitTestVisible;
            Shape.IsHitTestVisible = true;

            // HitTest receives logical coordinates, so because we need to talk to the outside world (WPF)
            // we need to convert to WPF's physical coordinates (pixels) from our internal logical coordinates
            point = ToPhysical(point);

            IFigure result = null;

#if SILVERLIGHT
            // surprisingly, UIElement.HitTest expects the point to be in global coordinates
            // and not in the coordinates of its parent.
            // That's why we need to translate the argument point
            // from Canvas coordinates to global coordinates

            // get the transform that will convert Canvas coordinates to RootVisual coordinates
            var transform = Shape.TransformToVisual(Application.Current.RootVisual);
            // and apply it to the argument point
            point = Shape.RenderTransform.Inverse.Transform(point);  // Take into account RenderTransform. - D.H.
            var hitTestPoint = transform.Transform(point);

            // finally, call HitTest with the point in global coordinates
            result = VisualTreeHelper.FindElementsInHostCoordinates(hitTestPoint, Shape).Any() ? this : null;
#else
            result = Avalonia.Input.InputExtensions.InputHitTest(Shape, point) != null ? this : null;
#endif
            Shape.IsHitTestVisible = oldHitTestVisible;

            return result;
        }

        public override bool Visible
        {
            get
            {
                return mVisible;
            }
            set
            {
                mVisible = value;
                UpdateShapeVisibility();
            }
        }

        public override bool Enabled
        {
            get
            {
                return mEnabled;
            }
            set
            {
                mEnabled = value;
                UpdateShapeAppearance();
            }
        }

        public override bool Selected
        {
            get
            {
                return base.Selected;
            }
            set
            {
                base.Selected = value;
                UpdateShapeAppearance();
                UpdateShapeVisibility();
            }
        }

        public override bool Locked
        {
            get
            {
                return base.Locked;
            }
            set
            {
                base.Locked = value;
                UpdateShapeAppearance();
            }
        }

        /// <summary>How much of a hidden figure shows while it is selected (<see cref="IsGhost"/>)</summary>
        public const double GhostOpacity = 0.5;

        bool ghost;
        bool hitTestVisibleBeforeGhost;

        /// <summary>
        /// A hidden figure that is selected (picked in the Figure List, say) shows faintly, so
        /// that one sees what was picked, and goes again when it is unselected. It stays hidden to
        /// everything else: nothing hits it, nothing snaps to it or drags it, since all of that
        /// asks <see cref="Visible"/>, which it isn't. Only a figure of the drawing itself: the
        /// parts of a composite follow their own rules (a vector's segment is never shown).
        /// </summary>
        protected bool IsGhost
        {
            get
            {
                return ghost;
            }
        }

        /// <summary>Whether the shape is on the canvas and wants its geometry kept up to date</summary>
        protected bool IsShown
        {
            get
            {
                return Exists && (Visible || ghost);
            }
        }

        protected void UpdateShapeVisibility()
        {
            if (Shape == null)
            {
                return;
            }

            bool wasGhost = ghost;
            ghost = !Visible
                && Selected
                && Exists
                && Drawing != null
                && Drawing.Figures.Contains(this);

            bool needsToBeVisible = IsShown;
            if ((Shape.Visibility == Visibility.Visible) != needsToBeVisible)
            {
                Shape.Visibility = needsToBeVisible ? Visibility.Visible : Visibility.Collapsed;
            }

            if (ghost == wasGhost)
            {
                return;
            }

            // a control (a check box, a link) of a label mustn't take clicks either
            if (ghost)
            {
                hitTestVisibleBeforeGhost = Shape.IsHitTestVisible;
                Shape.IsHitTestVisible = false;
                Shape.Opacity = GhostOpacity;

                // a hidden figure isn't kept in place while it is hidden
                if (Drawing != null && Drawing.Canvas != null)
                {
                    UpdateVisual();
                }
            }
            else
            {
                Shape.IsHitTestVisible = hitTestVisibleBeforeGhost;
                Shape.Opacity = 1;
            }
        }

        public override void ApplyStyle()
        {
            if (this.Style == null)
            {
                return;
            }
            this.Apply(Shape, Style);
            if (Drawing != null)
            {
                UpdateVisual();
            }
        }

        protected virtual void UpdateShapeAppearance()
        {
            ApplyStyle();
        }

        public override void OnAddingToCanvas(Canvas newContainer)
        {
            base.OnAddingToCanvas(newContainer);
            newContainer.Children.Add(Shape);
        }

        public override void OnRemovingFromCanvas(Canvas leavingContainer)
        {
            base.OnRemovingFromCanvas(leavingContainer);
            leavingContainer.Children.Remove(Shape);
        }
    }
}
