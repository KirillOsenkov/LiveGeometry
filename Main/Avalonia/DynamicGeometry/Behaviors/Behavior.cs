using System;
using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using System.Linq;

namespace DynamicGeometry
{
    public abstract partial class Behavior : INotifyPropertyChanged
    {
        public static Behavior Default { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void RaisePropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public virtual string Category
        {
            get
            {
                return "Common";
            }
        }

        public abstract string Name
        {
            get;
        }

        public virtual string HintText
        {
            get
            {
                return "";
            }
        }

        public virtual string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
        {
            string expectedFigure = "";
            if (args.FigureTypeNeeded.HasInterface<IPoint>())
            {
                expectedFigure = "point";
            }
            else if (args.FigureTypeNeeded == typeof(Vector))
            {
                expectedFigure = "vector or enter values";
            }
            else if (args.FigureTypeNeeded == typeof(IAngleProvider))
            {
                expectedFigure = "figure with an angle such as an arc or enter value";
            }
            else if (args.FigureTypeNeeded == typeof(ILengthProvider))
            {
                expectedFigure = "figure with length such as a segment or enter value";
            }
            else if (args.FigureTypeNeeded.HasInterface<ILine>())
            {
                expectedFigure = "line, ray or a segment";
            }
            else if (args.FigureTypeNeeded.HasInterface<ICircle>())
            {
                expectedFigure = "circle";
            }
            else if (args.FigureTypeNeeded.HasInterface<IEllipse>())
            {
                expectedFigure = "circle or ellipse";
            }
            else if (args.FigureTypeNeeded.HasInterface<ILinearFigure>())
            {
                expectedFigure = "line or a circle";
            }
            else
            {
                expectedFigure = args.FigureTypeNeeded.Name;
            }
            string hint = string.Format("Select a {0}.", expectedFigure);
            return hint;
        }

        private UIElement icon;
        public UIElement Icon
        {
            get
            {
                if (icon == null)
                {
                    icon = CreateIcon();
                }
                return icon;
            }
        }

        public abstract FrameworkElement CreateIcon();

        protected Drawing mDrawing;
        public virtual Drawing Drawing
        {
            get
            {
                return mDrawing;
            }
            set
            {
                if (mDrawing != null)
                {
                    mDrawing.OnAttachToCanvas -= mDrawing_OnAttachToCanvas;
                    mDrawing.OnDetachFromCanvas -= mDrawing_OnDetachFromCanvas;
                    ParentCanvas = null;
                }
                mDrawing = value;
                if (mDrawing != null)
                {
                    mDrawing.OnAttachToCanvas += mDrawing_OnAttachToCanvas;
                    mDrawing.OnDetachFromCanvas += mDrawing_OnDetachFromCanvas;
                    ParentCanvas = mDrawing.Canvas;
                }
            }
        }

        void mDrawing_OnAttachToCanvas(Canvas e)
        {
            ParentCanvas = e;
        }

        void mDrawing_OnDetachFromCanvas(Canvas e)
        {
            ParentCanvas = null;
        }

        private Canvas mParentCanvas;
        public virtual Canvas ParentCanvas
        {
            get
            {
                return mParentCanvas;
            }
            set
            {
                if (mParentCanvas != null)
                {
                    clickPreview.Clear();
                    mParentCanvas.PointerExited -= PointerExitedHandler;
                    mParentCanvas.PointerWheelChanged -= PointerWheelHandler;
                    mParentCanvas.PointerPressed -= PointerPressedHandler;
                    mParentCanvas.PointerMoved -= PointerMovedHandler;
                    mParentCanvas.PointerReleased -= PointerReleasedHandler;
                    mParentCanvas.KeyDown -= SafeKeyDown;
                    mParentCanvas.KeyUp -= SafeKeyUp;
                    mParentCanvas.Cursor = null;
                }
                mParentCanvas = value;
                if (mParentCanvas != null)
                {
                    mParentCanvas.PointerExited += PointerExitedHandler;
                    mParentCanvas.PointerWheelChanged += PointerWheelHandler;
                    mParentCanvas.PointerPressed += PointerPressedHandler;
                    mParentCanvas.PointerMoved += PointerMovedHandler;
                    mParentCanvas.PointerReleased += PointerReleasedHandler;
                    mParentCanvas.KeyDown += SafeKeyDown;
                    mParentCanvas.KeyUp += SafeKeyUp;
                }
            }
        }

        public virtual object PropertyBag
        {
            get
            {
                return null;
            }
        }

        public virtual void Started()
        {

        }

        public virtual void Stopping()
        {

        }

        public void Restart()
        {
            Stopping();
            Started();
        }

        public virtual bool IsInInitialState 
        {
            get
            {
                return true;
            }
        }

#if !PLAYER

        protected void AbortAndSetDefaultTool()
        {
            Drawing.SetDefaultBehavior();
        }

#endif

        // Avalonia has no global Keyboard.Modifiers; every input event carries the
        // modifier state, so track the last observed state as events flow through.
        static KeyModifiers currentModifiers;

        public static bool IsCtrlPressed()
        {
            return (currentModifiers & KeyModifiers.Control) == KeyModifiers.Control;
        }

        // Adapters translating Avalonia pointer events onto the WPF-shaped
        // MouseDown/MouseMove/MouseUp/MouseWheel virtuals that behaviors override.
        public static bool IsShiftPressed()
        {
            return (currentModifiers & KeyModifiers.Shift) == KeyModifiers.Shift;
        }

        void PointerPressedHandler(object sender, PointerPressedEventArgs e)
        {
            currentModifiers = e.KeyModifiers;
            clickPreview.Clear();
            var properties = e.GetCurrentPoint(mParentCanvas).Properties;
            if (properties.IsLeftButtonPressed)
            {
                SafeMouseDown(sender, e);
            }
            else if (properties.IsRightButtonPressed)
            {
                try
                {
                    MouseRightClick(sender, e);
                }
                catch (Exception ex)
                {
                    HandleException(ex);
                }
            }
        }

        void PointerMovedHandler(object sender, PointerEventArgs e)
        {
            currentModifiers = e.KeyModifiers;
            SafeMouseMove(sender, e);
            if (!errorHappened && !e.GetCurrentPoint(mParentCanvas).Properties.IsLeftButtonPressed)
            {
                UpdateCursor(e);
            }

            UpdateClickPreview(e);
        }

        void PointerExitedHandler(object sender, PointerEventArgs e)
        {
            clickPreview.Clear();
        }

        #region Click preview

        readonly ClickPreview clickPreview = new ClickPreview();

        void UpdateClickPreview(PointerEventArgs e)
        {
            if (errorHappened || mParentCanvas == null || Drawing == null)
            {
                clickPreview.Clear();
                return;
            }

            try
            {
                var placement = GetClickPreview(e);
                clickPreview.Show(
                    Drawing,
                    placement,
                    GetFigureToPick(e),
                    placement != null ? GetClickPreviewPointStyle(placement) : null);
            }
            catch (Exception)
            {
                clickPreview.Clear();
            }
        }

        /// <summary>
        /// What a click here would do, if it is a point worth announcing: one on a figure, at an
        /// intersection or in the middle of a segment. Null for anything else.
        /// </summary>
        protected virtual PointPlacement GetClickPreview(MouseEventArgs e)
        {
            return null;
        }

        /// <summary>
        /// The figure (not a point) a click here would pick for the tool, e.g. the line to be
        /// perpendicular to. Null if there is none.
        /// </summary>
        protected virtual IFigure GetFigureToPick(MouseEventArgs e)
        {
            return null;
        }

        /// <summary>
        /// The style the previewed point is going to get
        /// </summary>
        protected virtual IFigureStyle GetClickPreviewPointStyle(PointPlacement placement)
        {
            string name = StyleManager.FreePointStyleName;
            switch (placement.Kind)
            {
                case PointPlacementKind.OnFigure:
                    name = StyleManager.PointOnFigureStyleName;
                    break;
                case PointPlacementKind.Intersection:
                    name = StyleManager.IntersectionPointStyleName;
                    break;
                case PointPlacementKind.Midpoint:
                    name = StyleManager.MidpointStyleName;
                    break;
            }

            // a drawing from an older file has no styles by kind
            return Drawing.StyleManager.GetStyle(name)
                ?? Drawing.StyleManager.GetStyles<PointStyle>().FirstOrDefault();
        }

        #endregion

        /// <summary>
        /// Like in the original DG: a right-click gets you out of whatever you're doing.
        /// In the middle of a construction it cancels the construction, otherwise it
        /// switches back to the default (drag) tool.
        /// </summary>
        public virtual void MouseRightClick(object sender, MouseButtonEventArgs e)
        {
#if !PLAYER
            if (IsInInitialState)
            {
                AbortAndSetDefaultTool();
            }
            else
            {
                Restart();
            }
#endif
        }

        #region Cursor

        protected static readonly Cursor ArrowCursor = new Cursor(StandardCursorType.Arrow);
        protected static readonly Cursor CrossCursor = new Cursor(StandardCursorType.Cross);
        protected static readonly Cursor HandCursor = new Cursor(StandardCursorType.Hand);
        protected static readonly Cursor MoveCursor = new Cursor(StandardCursorType.SizeAll);
        protected static readonly Cursor NoCursor = new Cursor(StandardCursorType.No);

        void UpdateCursor(PointerEventArgs e)
        {
            if (mParentCanvas == null || Drawing == null)
            {
                return;
            }

            Cursor cursor;
            try
            {
                cursor = GetCursor(Coordinates(e, false, false, false));
            }
            catch (Exception)
            {
                cursor = ArrowCursor;
            }

            if (mParentCanvas.Cursor != cursor)
            {
                mParentCanvas.Cursor = cursor;
            }
        }

        /// <summary>
        /// The cursor tells what a click at this place would do:
        /// a cross - a new free point appears here;
        /// a hand - the click picks something that is already there, a figure or a place
        /// defined by figures (an intersection, a midpoint);
        /// an arrow - everything else, including a new point that slides along a figure.
        /// </summary>
        /// <param name="coordinates">Logical coordinates under the mouse</param>
        protected virtual Cursor GetCursor(Point coordinates)
        {
            return ArrowCursor;
        }

        protected static Cursor GetCursor(PointPlacement placement)
        {
            if (placement == null)
            {
                return ArrowCursor;
            }

            switch (placement.Kind)
            {
                case PointPlacementKind.Free:
                    return CrossCursor;
                case PointPlacementKind.Existing:
                case PointPlacementKind.Intersection:
                case PointPlacementKind.Midpoint:
                    return HandCursor;
                default:
                    return ArrowCursor;
            }
        }

        #endregion

        void PointerReleasedHandler(object sender, PointerReleasedEventArgs e)
        {
            currentModifiers = e.KeyModifiers;
            if (e.InitialPressMouseButton == MouseButton.Left)
            {
                SafeMouseUp(sender, e);
            }
        }

        void PointerWheelHandler(object sender, PointerWheelEventArgs e)
        {
            currentModifiers = e.KeyModifiers;
            clickPreview.Clear();
            MouseWheel(sender, e);
        }

        public virtual void KeyDown(object sender, KeyEventArgs e)
        {
#if !PLAYER
            if (e.Key == Avalonia.Input.Key.Escape)
            {
                AbortAndSetDefaultTool();
                e.Handled = true;
            }

            else if (e.Key == Avalonia.Input.Key.Delete)
            {
                try
                {
                    Drawing.DeleteSelection();
                }
                catch (Exception)
                {
                }
            }
            else if (e.Key == Avalonia.Input.Key.A)
            {
                Settings.Instance.AutoLabelPoints = !Settings.Instance.AutoLabelPoints;
                string state = (Settings.Instance.AutoLabelPoints) ? "on." : "off.";
                Drawing.RaiseStatusNotification("Toggling automatic labeling of points " + state);
            }
            else if (e.Key == Avalonia.Input.Key.G)
            {
                Settings.Instance.EnableSnapToGrid = !Settings.Instance.EnableSnapToGrid;
                string state = (Settings.Instance.EnableSnapToGrid) ? "on." : "off.";
                Drawing.RaiseStatusNotification("Toggling constraining to grid " + state);
            }

#endif
        }

        public virtual void KeyUp(object sender, KeyEventArgs e)
        {
        }

        public virtual void MouseDown(object sender, MouseButtonEventArgs e)
        {
        }

        public virtual void MouseMove(object sender, MouseEventArgs e)
        {
        }

        public virtual void MouseUp(object sender, MouseButtonEventArgs e)
        {
        }

        void HandleException(Exception ex)
        {
            Drawing.RaiseError(this, ex);
        }

        protected bool errorHappened = false;
        void SafeKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                KeyDown(sender, e);
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        void SafeKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                KeyUp(sender, e);
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        void SafeMouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                MouseDown(sender, e);
            }
            catch (Exception ex)
            {
                errorHappened = true;
                HandleException(ex);
            }
        }

        void SafeMouseMove(object sender, MouseEventArgs e)
        {
            if (errorHappened)
            {
                return;
            }
            try
            {
                MouseMove(sender, e);
            }
            catch (Exception ex)
            {
                HandleException(ex);
                errorHappened = true;
            }
        }

        void SafeMouseUp(object sender, MouseButtonEventArgs e)
        {
            errorHappened = false;
            try
            {
                MouseUp(sender, e);
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        const double WheelZoomFactor = 1.2;

        public virtual void MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Drawing != null)
            {
                // around the cursor; a touchpad sends many small deltas, a wheel notch is 1
                if (e.Delta.Y != 0)
                {
                    var factor = System.Math.Pow(WheelZoomFactor, e.Delta.Y);
                    Drawing.CoordinateSystem.Zoom(factor, e.GetPosition(ParentCanvas));
                }
            }
        }

        #region Coordinates

        protected virtual Point Coordinates(MouseEventArgs e)
        {
            // Like in the original DG, holding Shift snaps to the grid without having to
            // turn the setting on.
            if (IsShiftPressed() && !Settings.Instance.EnableSnapToGrid)
            {
                return Coordinates(e, false, true, false);
            }

            return Coordinates(e, Settings.Instance.EnableSnapToPoint, Settings.Instance.EnableSnapToGrid, Settings.Instance.EnableSnapToCenter);
        }

        protected virtual Point Coordinates(MouseEventArgs e, bool snapToPoint, bool snapToGrid, bool snapToCenter)
        {
            var result = e.GetPosition(ParentCanvas);
            result = ToLogical(result);

            if (snapToCenter)
            {
                result = Math.GetSnapToSegmentCenterPosition(result, new List<Segment>(Drawing.Figures.Where(f => f.Visible).ToSegments(result)));
            }
            else
            {
                if (snapToPoint)
                {
                    result = Math.GetSnapToPointPosition(Drawing.CoordinateSystem.MajorGridStep, result, new List<Point>(Drawing.Figures.Where(f => f.Visible).ToPoints()), Settings.Instance.EnableSnapToGrid);
                }
                else if (snapToGrid)
                {
                    // the labeled lines, whatever the zoom made of them
                    result = Math.GetSnapToGridPosition(Drawing.CoordinateSystem.MajorGridStep, result);
                }
            }
            return result;
        }

        protected double CursorTolerance
        {
            get
            {
                return Drawing.CoordinateSystem.CursorTolerance;
            }
        }

        protected double ToPhysical(double logicalLength)
        {
            return Drawing.CoordinateSystem.ToPhysical(logicalLength);
        }

        protected Point ToPhysical(Point point)
        {
            return Drawing.CoordinateSystem.ToPhysical(point);
        }

        protected double ToLogical(double pixelLength)
        {
            return Drawing.CoordinateSystem.ToLogical(pixelLength);
        }

        protected Point ToLogical(Point pixel)
        {
            return Drawing.CoordinateSystem.ToLogical(pixel);
        }

        protected Avalonia.Rect ToLogical(Avalonia.Rect rect)
        {
            return Drawing.CoordinateSystem.ToLogical(rect);
        }

        protected Avalonia.Rect ToPhysical(Avalonia.Rect rect)
        {
            return Drawing.CoordinateSystem.ToPhysical(rect);
        }
        #endregion
    }
}
