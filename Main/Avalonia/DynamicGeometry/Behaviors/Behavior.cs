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
                    mParentCanvas.PointerCaptureLost -= PointerCaptureLostHandler;
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
                    mParentCanvas.PointerCaptureLost += PointerCaptureLostHandler;
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

        public static bool IsAltPressed()
        {
            return (currentModifiers & KeyModifiers.Alt) == KeyModifiers.Alt;
        }

        void PointerPressedHandler(object sender, PointerPressedEventArgs e)
        {
            currentModifiers = e.KeyModifiers;
            clickPreview.Clear();
            if (e.Pointer.Type == PointerType.Touch)
            {
                TouchPressed(sender, e);
                return;
            }

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
            if (e.Pointer.Type == PointerType.Touch)
            {
                TouchMoved(sender, e);
                return;
            }

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

        #region Touch

        // Fingers. A mouse press acts at once (a point appears under the button going
        // down), but the first finger of two can't: it would leave a point, or start a drag,
        // before the second one says that the two are zooming. So a finger's press is kept
        // back until it is known for what it is - a tap when the finger lifts where it
        // came down, a drag once it has gone a little way (the tool then gets the press
        // where the finger came down, and the moves), or one of two fingers, which zoom and
        // pan the view and of which the tool hears nothing. Unhandled, a second finger was a
        // second press to the tool: with the Drag tool the view jumped between the two
        // fingers on every move, with any other a pinch left points behind.

        class Finger
        {
            public IPointer Pointer;
            public Point Position;
        }

        // The fingers on the canvas in the order they came down, each where it was last
        // seen, in canvas pixels. Static: a finger outlives a change of tool.
        static readonly List<Finger> fingers = new List<Finger>();

        // from the second finger down until the last one is up
        static bool isPinching;

        // the one finger on the canvas: its press, the tool it is for, whether the tool
        // has had it, and the last that was seen of the finger
        static PointerPressedEventArgs touchPress;
        static Behavior touchTool;
        static bool touchPressDelivered;
        static PointerEventArgs touchLast;

        /// <summary>In pixels: how far a finger goes from where it came down before it is dragging and not tapping</summary>
        public static double TouchSlop = 6;

        /// <summary>
        /// In pixels: what a finger reaches, where the mouse has
        /// <see cref="Settings.CursorTolerance"/> - in force only while a tool handles what
        /// a finger did
        /// </summary>
        public static double TouchTolerance = 10;

        void AsTouch(Action action)
        {
            double tolerance = Math.CursorTolerance;
            Math.CursorTolerance = TouchTolerance;
            try
            {
                action();
            }
            finally
            {
                Math.CursorTolerance = tolerance;
            }
        }

        static void ForgetTouch()
        {
            touchPress = null;
            touchTool = null;
            touchPressDelivered = false;
            touchLast = null;
        }

        void TouchPressed(object sender, PointerPressedEventArgs e)
        {
            // a finger whose release never arrived holds nothing any more
            fingers.RemoveAll(finger => finger.Pointer == e.Pointer || finger.Pointer.Captured == null);
            fingers.Add(new Finger() { Pointer = e.Pointer, Position = e.GetPosition(mParentCanvas) });
            if (fingers.Count == 1)
            {
                isPinching = false;
                touchPress = e;
                touchTool = this;
                touchPressDelivered = false;
                touchLast = e;
                return;
            }

            if (!isPinching)
            {
                isPinching = true;

                // a drag under way ends where the first finger is
                if (touchPressDelivered && touchTool == this && touchLast != null)
                {
                    var last = touchLast;
                    AsTouch(() => SafeMouseUp(sender, last));
                }

                ForgetTouch();
            }
        }

        void TouchMoved(object sender, PointerEventArgs e)
        {
            int index = fingers.FindIndex(finger => finger.Pointer == e.Pointer);
            if (index < 0)
            {
                return;
            }

            var position = e.GetPosition(mParentCanvas);
            if (isPinching)
            {
                if (index < 2 && fingers.Count >= 2 && Drawing != null)
                {
                    var from = Math.Midpoint(fingers[0].Position, fingers[1].Position);
                    double span = fingers[0].Position.Distance(fingers[1].Position);
                    fingers[index].Position = position;
                    var to = Math.Midpoint(fingers[0].Position, fingers[1].Position);
                    double newSpan = fingers[0].Position.Distance(fingers[1].Position);

                    // fingers almost on each other say nothing about the zoom
                    double factor = span > TouchSlop && newSpan > TouchSlop ? newSpan / span : 1;
                    Drawing.CoordinateSystem.PanAndZoom(from, to, factor);
                }
                else
                {
                    fingers[index].Position = position;
                }

                return;
            }

            fingers[index].Position = position;
            if (touchTool != this || touchPress == null)
            {
                return;
            }

            touchLast = e;
            if (!touchPressDelivered)
            {
                if (position.Distance(touchPress.GetPosition(mParentCanvas)) < TouchSlop)
                {
                    return;
                }

                touchPressDelivered = true;
                var press = touchPress;
                AsTouch(() => SafeMouseDown(sender, press));
            }

            AsTouch(() => SafeMouseMove(sender, e));
            UpdateClickPreview(e);
        }

        /// <param name="e">The release; null when the touch was taken away, and the finger is where it was last seen</param>
        void TouchReleased(object sender, IPointer pointer, PointerEventArgs e)
        {
            int index = fingers.FindIndex(finger => finger.Pointer == pointer);
            if (index < 0)
            {
                return;
            }

            fingers.RemoveAt(index);
            clickPreview.Clear();
            if (isPinching)
            {
                if (fingers.Count == 0)
                {
                    isPinching = false;
                }

                return;
            }

            if (touchTool == this && touchPress != null)
            {
                if (e != null && !touchPressDelivered)
                {
                    // a tap: the press and the release in one go
                    touchPressDelivered = true;
                    var press = touchPress;
                    AsTouch(() => SafeMouseDown(sender, press));
                }

                var release = e ?? touchLast;
                if (touchPressDelivered && release != null)
                {
                    AsTouch(() => SafeMouseUp(sender, release));
                }
            }

            ForgetTouch();
        }

        // A touch the system took away (a gesture of its own, the window going away): a
        // drag ends where it is, a press that was not one yet is nothing. After a release
        // the finger is already gone from the list.
        void PointerCaptureLostHandler(object sender, PointerCaptureLostEventArgs e)
        {
            if (e.Pointer.Type == PointerType.Touch)
            {
                TouchReleased(sender, e.Pointer, e: null);
            }
        }

        #endregion

        void PointerReleasedHandler(object sender, PointerReleasedEventArgs e)
        {
            currentModifiers = e.KeyModifiers;
            if (e.Pointer.Type == PointerType.Touch)
            {
                TouchReleased(sender, e.Pointer, e);
                return;
            }

            if (e.InitialPressMouseButton == MouseButton.Left)
            {
                SafeMouseUp(sender, e);
            }

            // what the press did may have changed what a click here does (a dragged point
            // dropped onto a figure no longer shows the halo of its snap)
            UpdateClickPreview(e);
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

            // (A and G used to toggle "Label new points" and "Snap to grid" here, for the
            // tools that don't handle keys themselves - Point, Slider, Text. They are tool
            // letters now, A for the arc and G for the grid: with the Point tool on, G
            // showed the grid and, without a word, made every new point snap to it.)

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
