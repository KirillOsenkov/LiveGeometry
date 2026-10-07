using Avalonia;
using Avalonia.Controls;
using Avalonia.VisualTree;

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
                if (NameLabel != null && Drawing != null)
                {
                    NameLabel.UpdateVisual();
                }
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
            if (needsToBeVisible && canvas != null && !isShapeOnCanvas)
            {
                // held back while hidden (OnAddingToCanvas): on the canvas now, with the
                // geometry it went without (UpdateVisual overrides skip a hidden figure)
                PutShapeOnCanvas();
                UpdateVisual();
            }

            if ((Shape.Visibility == Visibility.Visible) != needsToBeVisible)
            {
                Shape.Visibility = needsToBeVisible ? Visibility.Visible : Visibility.Collapsed;
            }

            UpdateSelectionHalo();

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

        // The canvas the figure is on, and whether its shape is among the canvas's children.
        // A figure hidden when it comes onto the canvas keeps its shape off it until it first
        // shows (Visible, or the ghost of a selection): the hidden helpers of a construction
        // and the variables of a generated drawing - most of what Stretchy Slime is made of -
        // then cost the canvas nothing, at the load and at every frame. The shape is styled
        // all the same (code reads a point's size or a line's thickness off it, hidden or
        // not). Only for shapes proper: a label is a control, measured as shown while it is
        // hidden (its room is kept), which wants it in the tree, where it inherits the
        // window's font.
        Canvas canvas;
        bool isShapeOnCanvas;

        bool DefersHiddenShape => Shape is Avalonia.Controls.Shapes.Shape;

        void PutShapeOnCanvas()
        {
            if (!isShapeOnCanvas && canvas != null)
            {
                isShapeOnCanvas = true;
                canvas.Children.Add(Shape);
            }
        }

        public override void OnAddingToCanvas(Canvas newContainer)
        {
            // the style first (the base assigns one), onto the canvas second: a property set on
            // a shape in the tree invalidates it, on one outside it is only a set
            base.OnAddingToCanvas(newContainer);
            canvas = newContainer;
            if (!DefersHiddenShape || Visible || Selected)
            {
                PutShapeOnCanvas();
            }

            UpdateSelectionHalo();
        }

        public override void OnRemovingFromCanvas(Canvas leavingContainer)
        {
            base.OnRemovingFromCanvas(leavingContainer);
            if (isShapeOnCanvas)
            {
                leavingContainer.Children.Remove(Shape);
                isShapeOnCanvas = false;
            }

            canvas = null;
            UpdateSelectionHalo();
        }

        SelectionHalo selectionHalo;

        /// <summary>
        /// Whether a selection shows a <see cref="SelectionHalo"/> along the shape: not for a
        /// part that is never drawn (a vector's segment inside its arrow)
        /// </summary>
        public bool ShowsSelectionHalo { get; set; } = true;

        void UpdateSelectionHalo()
        {
            var geometryShape = Shape as Avalonia.Controls.Shapes.Shape;
            var canvas = geometryShape?.GetVisualParent() as Canvas;
            bool needed = Selected
                && IsShown
                && ShowsSelectionHalo
                && canvas != null;
            if (needed == (selectionHalo != null))
            {
                return;
            }

            if (needed)
            {
                selectionHalo = new SelectionHalo(geometryShape, isPoint: this is PointBase);
                selectionHalo.AddTo(canvas);
            }
            else
            {
                selectionHalo.Remove();
                selectionHalo = null;
            }
        }
    }
}
