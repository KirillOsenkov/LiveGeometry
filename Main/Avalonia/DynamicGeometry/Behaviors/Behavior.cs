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
            return "Select " + DescribeFigureType(args.FigureTypeNeeded) + ".";
        }

        /// <summary>A figure of the type a tool asks for, in words: "a point", "a line, ray or a segment", "a slider"</summary>
        public static string DescribeFigureType(Type type)
        {
            if (type.HasInterface<IPoint>())
            {
                return "a point";
            }

            if (type == typeof(Vector))
            {
                return "a vector or enter values";
            }

            if (type == typeof(IAngleProvider))
            {
                return "a figure with an angle such as an arc or enter value";
            }

            if (type == typeof(ILengthProvider))
            {
                return "a figure with length such as a segment or enter value";
            }

            if (type.HasInterface<ILine>())
            {
                return "a line, ray or a segment";
            }

            if (type.HasInterface<ICircle>())
            {
                return "a circle";
            }

            if (type.HasInterface<IEllipse>())
            {
                return "a circle or ellipse";
            }

            if (type.HasInterface<ILinearFigure>())
            {
                return "a line or a circle";
            }

            // the type's name in words: AngleMeasurement is "an angle measurement"
            var words = System.Text.RegularExpressions.Regex.Replace(type.Name, "(?<=[a-z])(?=[A-Z])", " ").ToLowerInvariant();
            return ("aeiou".IndexOf(words[0]) >= 0 ? "an " : "a ") + words;
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
                    ForgetChoice();
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
                    // a move that waits for a frame was the tool's that is put down, and so
                    // was a press of the right button
                    ForgetWaitingMove();
                    ForgetRightPress();
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

        protected void AbortAndSetDefaultTool()
        {
            Drawing.SetDefaultBehavior();
        }

        // Avalonia has no global Keyboard.Modifiers; every input event carries the
        // modifier state, so track the last observed state as events flow through.
        static KeyModifiers currentModifiers;

        /// <summary>Ctrl, or Cmd on a Mac (where Ctrl+click is a right click)</summary>
        public static bool IsCtrlPressed()
        {
            return (currentModifiers & (KeyModifiers.Control | KeyModifiers.Meta)) != 0;
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

        /// <summary>
        /// Ctrl+click is a right click on a Mac. AppKit hands it over as a press of the left
        /// button with Control held, and Avalonia passes it on as such: on the desktop it was
        /// a click (with Ctrl, which stands for Cmd) where a Mac user asks for the context menu.
        /// The moves and the release of that press are no left button's either.
        /// </summary>
        bool isSecondaryClickHeld;

        static bool IsSecondaryClick(PointerPointProperties properties, KeyModifiers modifiers)
        {
            return KeyNames.IsMac && properties.IsLeftButtonPressed && (modifiers & KeyModifiers.Control) != 0;
        }

        void PointerPressedHandler(object sender, PointerPressedEventArgs e)
        {
            StartOverMoves();
            currentModifiers = e.KeyModifiers;
            clickPreview.Clear();
            if (e.Pointer.Type == PointerType.Touch)
            {
                TouchPressed(sender, e);
                return;
            }

            var properties = e.GetCurrentPoint(mParentCanvas).Properties;
            isSecondaryClickHeld = IsSecondaryClick(properties, e.KeyModifiers);
            if (properties.IsLeftButtonPressed && !isSecondaryClickHeld)
            {
                SafeMouseDown(sender, e);

                // a choice is for one click: the next step starts at its first option
                Choice.Forget();
                OfferClickOptions(e);
            }
            else if (properties.IsRightButtonPressed || isSecondaryClickHeld)
            {
                PressRightButton(e);
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

            if (movedThisFrame)
            {
                waitingMove = e;
                waitingMoveSender = sender;
                return;
            }

            movedThisFrame = WaitForFrame();
            HandleMove(sender, e);
        }

        void HandleMove(object sender, PointerEventArgs e)
        {
            if (MoveRightButton(sender, e))
            {
                return;
            }

            lastMove = e;
            lastMoveSender = sender;
            OfferClickOptions(e);
            SafeMouseMove(sender, e);
            if (!errorHappened && !e.GetCurrentPoint(mParentCanvas).Properties.IsLeftButtonPressed)
            {
                UpdateCursor(e);
            }

            UpdateClickPreview(e);
        }

        void PointerExitedHandler(object sender, PointerEventArgs e)
        {
            HandleWaitingMove();
            clickPreview.Clear();
            ForgetChoice();
        }

        #region Choosing among overlapping figures

        /// <summary>Which of the things a click at the cursor could mean it means (<see cref="ClickChoice"/>)</summary>
        public ClickChoice Choice { get; protected set; } = new ClickChoice();

        // the last move over the canvas, to show anew what a click takes once Tab chose another
        PointerEventArgs lastMove;
        object lastMoveSender;

        // what the status says of the choice, null while there is none
        string shownChoiceStatus;

        /// <summary>
        /// What a click here could take, best first: the figures the step wants, or the points
        /// it could make (<see cref="PointPlacement"/>). More than one is a choice the user
        /// makes with Tab, or in a menu after a tap. Nothing by default.
        /// </summary>
        protected virtual IReadOnlyList<object> FindClickOptions(MouseEventArgs e)
        {
            return Array.Empty<object>();
        }

        IReadOnlyList<object> SafeFindClickOptions(MouseEventArgs e)
        {
            if (errorHappened || mParentCanvas == null || Drawing == null)
            {
                return Array.Empty<object>();
            }

            try
            {
                return FindClickOptions(e);
            }
            catch (Exception)
            {
                return Array.Empty<object>();
            }
        }

        void OfferClickOptions(MouseEventArgs e)
        {
            Choice.Offer(SafeFindClickOptions(e));
            ShowChoiceStatus();
        }

        void ShowChoiceStatus()
        {
            var text = Choice.StatusText;
            if (text != shownChoiceStatus)
            {
                shownChoiceStatus = text;
                Drawing?.RaiseChoiceStatus(text);
            }
        }

        protected void ForgetChoice()
        {
            Choice.Forget();
            ShowChoiceStatus();
        }

        /// <summary>
        /// Tab (Shift+Tab: back): the next of the things a click at the cursor could mean, with
        /// its halo and the status saying which it is. False when there is no choice there, and
        /// Tab is left to do what it does.
        /// </summary>
        public bool StepChoice(bool backwards)
        {
            if (lastMove == null || mParentCanvas == null || !Choice.Step(backwards ? -1 : 1))
            {
                return false;
            }

            ShowChoiceStatus();

            // what the hover shows, worked out anew for the option chosen
            SafeMouseMove(lastMoveSender, lastMove);
            UpdateCursor(lastMove);
            UpdateClickPreview(lastMove);
            return true;
        }

        /// <summary>
        /// A tap where it could mean more than one thing: a menu says what each is, and the one
        /// picked is the tap. Dismissed, the tap is nothing.
        /// </summary>
        void AskWhatTapMeans(
            object sender,
            PointerPressedEventArgs press,
            PointerEventArgs release,
            IReadOnlyList<object> options)
        {
            var menu = new ContextMenu();
            foreach (var option in options)
            {
                var item = new MenuItem() { Header = ClickChoice.Describe(option) };
                item.Click += (s, args) =>
                {
                    Choice.Offer(options);
                    Choice.Choose(option);
                    AsTouch(() => SafeMouseDown(sender, press));
                    AsTouch(() => SafeMouseUp(sender, release));
                    ForgetChoice();
                };
                menu.Items.Add(item);
            }

            menu.Open(mParentCanvas);
        }

        #endregion

        #region One move a frame

        // A mouse reports far more often than the screen is drawn (500 or 1000 times a
        // second, against 60), and Windows hands a move over whenever the app asks for its
        // messages: a drag in a heavy drawing worked the drawing out again for every move, a
        // dozen times for each frame that showed one of them. The first move since the last
        // frame is handled at once, as before; the ones after it wait for the next frame,
        // which handles the last of them before it lays out and draws. A press, a release,
        // the wheel and the pointer leaving handle the waiting one first, so that everything
        // comes in order. (The browser hands over one move a frame already.) A drag of the
        // grid of the Inversion in a Circle at 1000 moves a second: of the 2000 moves that
        // came in 4 seconds, one a frame was handled, and the drag took a fifth of the time.

        PointerEventArgs waitingMove;
        object waitingMoveSender;
        bool movedThisFrame;
        bool frameRequested;

        /// <summary>Asks for the next frame, once; false where there is no frame to ask for (no window)</summary>
        bool WaitForFrame()
        {
            if (frameRequested)
            {
                return true;
            }

            var topLevel = TopLevel.GetTopLevel(mParentCanvas);
            if (topLevel == null)
            {
                return false;
            }

            frameRequested = true;
            topLevel.RequestAnimationFrame(_ => OnFrame());
            return true;
        }

        void OnFrame()
        {
            frameRequested = false;
            movedThisFrame = false;
            if (waitingMove != null && mParentCanvas != null)
            {
                movedThisFrame = WaitForFrame();
                HandleWaitingMove();
            }
        }

        /// <summary>
        /// Before a press or a release: the move that waits comes first, and the next one is
        /// handled at once again, as the first move of a drag should be (the status says what
        /// Shift and Alt do to the point as soon as it moves)
        /// </summary>
        void StartOverMoves()
        {
            HandleWaitingMove();
            movedThisFrame = false;
        }

        void HandleWaitingMove()
        {
            var e = waitingMove;
            var sender = waitingMoveSender;
            ForgetWaitingMove();
            if (e != null)
            {
                HandleMove(sender, e);
            }
        }

        void ForgetWaitingMove()
        {
            waitingMove = null;
            waitingMoveSender = null;
        }

        #endregion

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
                var angle = GetAngleToPick(e);
                if (angle != null)
                {
                    clickPreview.ShowAngle(Drawing, angle, PreviewsAngleBisector());
                    return;
                }

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

        /// <summary>An angle a click here would measure whole (the Angle tool near a vertex). Null if there is none.</summary>
        protected virtual AngleAtVertex GetAngleToPick(MouseEventArgs e)
        {
            return null;
        }

        /// <summary>Whether the preview of that angle also shows the bisector a click would make (the Angle Bisector tool)</summary>
        protected virtual bool PreviewsAngleBisector()
        {
            return false;
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
                case PointPlacementKind.InRegion:
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
            if (IsInInitialState)
            {
                AbortAndSetDefaultTool();
            }
            else
            {
                Restart();
            }
        }

        #region The right button

        // A right click (Ctrl+click on a Mac) is MouseRightClick at the release, when the
        // button comes up where it went down, as Windows opens a context menu. A drag with
        // the right button is the tool's, if it takes one (the Drag tool draws a lasso); a
        // tool that doesn't gets the click at the release wherever it is. The click came at
        // the press: the Drag tool's menu opened under the button still held, and a drag with
        // it could mean nothing.

        // the press while the button is held, null otherwise
        PointerPressedEventArgs rightPress;
        Point rightPressPosition;

        // whether the tool took the drag, and whether it said no to it
        bool isRightDragging;
        bool isRightDragRefused;

        void PressRightButton(PointerPressedEventArgs e)
        {
            rightPress = e;
            rightPressPosition = e.GetPosition(mParentCanvas);
            isRightDragging = false;
            isRightDragRefused = false;
        }

        /// <summary>The move of the right button held: a drag once it has gone far enough. False when it is not held.</summary>
        bool MoveRightButton(object sender, PointerEventArgs e)
        {
            if (rightPress == null)
            {
                return false;
            }

            var properties = e.GetCurrentPoint(mParentCanvas).Properties;
            if (!(isSecondaryClickHeld ? properties.IsLeftButtonPressed : properties.IsRightButtonPressed))
            {
                // its release never came (the window lost the mouse): a drag ends where it is
                ReleaseRightButton(sender, e, isReleased: false);
                return false;
            }

            if (!isRightDragging)
            {
                if (isRightDragRefused || e.GetPosition(mParentCanvas).Distance(rightPressPosition) < Dragger.DragThreshold)
                {
                    return true;
                }

                try
                {
                    isRightDragging = MouseRightDragStarted(sender, rightPress);
                }
                catch (Exception ex)
                {
                    HandleException(ex);
                }

                isRightDragRefused = !isRightDragging;
                if (!isRightDragging)
                {
                    return true;
                }
            }

            try
            {
                MouseRightDragMoved(sender, e);
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }

            return true;
        }

        /// <param name="isReleased">False when the release never came: no click then</param>
        void ReleaseRightButton(object sender, PointerEventArgs e, bool isReleased)
        {
            bool wasDragging = isRightDragging;
            ForgetRightPress();
            try
            {
                if (wasDragging)
                {
                    MouseRightDragEnded(sender, e);
                }
                else if (isReleased)
                {
                    MouseRightClick(sender, e);
                }
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        void ForgetRightPress()
        {
            rightPress = null;
            isRightDragging = false;
            isRightDragRefused = false;
            isSecondaryClickHeld = false;
        }

        /// <summary>
        /// A drag with the right button begins: it has gone <see cref="Dragger.DragThreshold"/>
        /// pixels from the press. True takes it: the moves and the release go to
        /// <see cref="MouseRightDragMoved"/> and <see cref="MouseRightDragEnded"/>, and there is
        /// no click. False by default: the release is a right click wherever it is.
        /// </summary>
        protected virtual bool MouseRightDragStarted(object sender, PointerEventArgs press)
        {
            return false;
        }

        protected virtual void MouseRightDragMoved(object sender, PointerEventArgs e)
        {
        }

        /// <summary>The release of a drag the tool took; also a move with the button no longer held, when the release never came</summary>
        protected virtual void MouseRightDragEnded(object sender, PointerEventArgs e)
        {
        }

        #endregion

        /// <summary>
        /// Escape: a gesture under way (a lasso being drawn) is dropped as if it never began,
        /// and nothing else happens. False when there is none, and Escape does what it does.
        /// </summary>
        public virtual bool AbandonGesture()
        {
            return false;
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
                // (a point inside a shape appears right under the cursor, as a free one does)
                case PointPlacementKind.Free:
                case PointPlacementKind.InRegion:
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
        // down), but a finger's can't: it would leave a point, or start a drag, before a
        // second finger says that the two are zooming. So a finger's press is kept back
        // until it is known for what it is - a tap when the finger lifts where it came
        // down, a drag once it has gone a little way (the tool then gets the press where
        // the finger came down, and the moves). Unhandled, a second finger was a second
        // press to the tool: with the Drag tool the view jumped between the two fingers on
        // every move, with any other a pinch left points behind.
        //
        // Each finger is for a figure or for the paper, decided where it comes down
        // (TakesFingerOnFigure: the Drag tool says a figure where a press would drag or
        // select one; a tool that constructs says the paper for every finger). A figure
        // finger is the tool's for as long as it is down: it taps or drags its figure, and
        // nothing another finger does takes it away. The paper fingers are the tool's only
        // while there is one finger on the canvas (a tap selects, a drag pans through the
        // tool, which records it); with any other finger they move the view themselves and
        // the tool hears nothing of them - one pans, two pan and zoom (the paper under them
        // stays under them), a third does nothing. So a thumb put down to zoom while a point
        // is held leaves the point in the finger that holds it, and two fingers drag the
        // two ends of a segment at once. Until 2026-10-11 the second finger always made a
        // pinch of the two, and ended the drag of the first.

        enum FingerRole
        {
            Paper,
            Figure
        }

        class Finger
        {
            public IPointer Pointer;

            // where it was last seen, in canvas pixels
            public Point Position;
            public FingerRole Role;

            // the tool it came down with
            public Behavior Tool;

            // its press, kept back until the finger is a tap or a drag; null once the
            // finger moves the view itself
            public PointerPressedEventArgs Press;
            public bool PressDelivered;
            public PointerEventArgs Last;

            // a paper finger that moves the view itself, past the tool
            public bool DrivesView;
        }

        // The fingers on the canvas in the order they came down. Static: a finger outlives
        // a change of tool.
        static readonly List<Finger> fingers = new List<Finger>();

        /// <summary>In pixels: how far a finger goes from where it came down before it is dragging and not tapping</summary>
        public static double TouchSlop = 6;

        /// <summary>
        /// In pixels: what a finger reaches, where the mouse has
        /// <see cref="Settings.CursorTolerance"/> - in force only while a tool handles what
        /// a finger did
        /// </summary>
        public static double TouchTolerance = 10;

        /// <summary>
        /// Whether a finger pressed here is for a figure - it taps or drags that figure,
        /// whatever other fingers do meanwhile - rather than for the paper, which the paper
        /// fingers pan and zoom together. False by default: a tool that constructs takes
        /// one finger at a time and lets a second one pinch. Asked with a finger's reach.
        /// </summary>
        protected virtual bool TakesFingerOnFigure(MouseButtonEventArgs e)
        {
            return false;
        }

        bool SafeTakesFingerOnFigure(MouseButtonEventArgs e)
        {
            if (errorHappened || mParentCanvas == null || Drawing == null)
            {
                return false;
            }

            try
            {
                return TakesFingerOnFigure(e);
            }
            catch (Exception)
            {
                return false;
            }
        }

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

        static Finger FindFinger(IPointer pointer)
        {
            return fingers.Find(finger => finger.Pointer == pointer);
        }

        void TouchPressed(object sender, PointerPressedEventArgs e)
        {
            // a finger whose release never arrived holds nothing any more
            fingers.RemoveAll(finger => finger.Pointer == e.Pointer || finger.Pointer.Captured == null);
            var finger = new Finger()
            {
                Pointer = e.Pointer,
                Position = e.GetPosition(mParentCanvas),
                Tool = this,
                Press = e,
                Last = e
            };
            bool onFigure = false;
            AsTouch(() => onFigure = SafeTakesFingerOnFigure(e));
            finger.Role = onFigure ? FingerRole.Figure : FingerRole.Paper;
            fingers.Add(finger);

            // the paper fingers are the tool's only while there is one finger
            if (fingers.Count > 1)
            {
                foreach (var paper in fingers.Where(other => other.Role == FingerRole.Paper && !other.DrivesView).ToList())
                {
                    TakeFromTool(sender, paper);
                }
            }

            // A figure finger's press goes to the tool at once (a press on a figure does
            // nothing visible under the Drag tool, which waits for the slop itself), unless
            // a tap here would have to ask what it means. Held back, the press reached the
            // tool after the paper fingers had moved the view, and found the figure no
            // longer under the finger.
            if (finger.Role == FingerRole.Figure)
            {
                IReadOnlyList<object> options = null;
                AsTouch(() => options = SafeFindClickOptions(e));
                if (options.Count <= 1)
                {
                    finger.PressDelivered = true;
                    AsTouch(() => SafeMouseDown(sender, e));
                }
            }
        }

        /// <summary>A paper finger goes over from the tool to the view: a drag under way ends where the finger is</summary>
        void TakeFromTool(object sender, Finger finger)
        {
            if (finger.PressDelivered && finger.Tool == this && finger.Last != null)
            {
                var last = finger.Last;
                AsTouch(() => SafeMouseUp(sender, last));
            }

            finger.Press = null;
            finger.PressDelivered = false;
            finger.Last = null;
            finger.DrivesView = true;
        }

        void TouchMoved(object sender, PointerEventArgs e)
        {
            var finger = FindFinger(e.Pointer);
            if (finger == null)
            {
                return;
            }

            var position = e.GetPosition(mParentCanvas);
            if (finger.DrivesView)
            {
                MoveView(sender, finger, position);
                return;
            }

            finger.Position = position;
            if (finger.Tool != this || finger.Press == null)
            {
                return;
            }

            finger.Last = e;
            if (!finger.PressDelivered)
            {
                if (position.Distance(finger.Press.GetPosition(mParentCanvas)) < TouchSlop)
                {
                    return;
                }

                finger.PressDelivered = true;
                var press = finger.Press;
                AsTouch(() => SafeMouseDown(sender, press));
            }

            AsTouch(() => SafeMouseMove(sender, e));
            UpdateClickPreview(e);
        }

        /// <summary>
        /// The paper fingers move the view: what was under the first two of them stays under
        /// them (one finger pans, two pan and zoom, a third is ignored). The figures held by
        /// the other fingers stay under those: each gets its last move again over the moved
        /// paper, or a held point drifted off its finger until the finger moved.
        /// </summary>
        void MoveView(object sender, Finger finger, Point position)
        {
            var driving = fingers.Where(other => other.DrivesView).Take(2).ToList();
            int index = driving.IndexOf(finger);
            if (index < 0 || Drawing == null)
            {
                finger.Position = position;
                return;
            }

            Point from;
            Point to;
            double factor = 1;
            if (driving.Count == 1)
            {
                from = finger.Position;
                finger.Position = position;
                to = position;
            }
            else
            {
                from = Math.Midpoint(driving[0].Position, driving[1].Position);
                double span = driving[0].Position.Distance(driving[1].Position);
                finger.Position = position;
                to = Math.Midpoint(driving[0].Position, driving[1].Position);
                double newSpan = driving[0].Position.Distance(driving[1].Position);

                // fingers almost on each other say nothing about the zoom
                factor = span > TouchSlop && newSpan > TouchSlop ? newSpan / span : 1;
            }

            Drawing.CoordinateSystem.PanAndZoom(from, to, factor);
            foreach (var held in fingers.Where(other => other.PressDelivered && other.Tool == this && other.Last != null).ToList())
            {
                var last = held.Last;
                AsTouch(() => SafeMouseMove(sender, last));
            }
        }

        /// <param name="e">The release; null when the touch was taken away, and the finger is where it was last seen</param>
        void TouchReleased(object sender, IPointer pointer, PointerEventArgs e)
        {
            var finger = FindFinger(pointer);
            if (finger == null)
            {
                return;
            }

            fingers.Remove(finger);
            clickPreview.Clear();
            if (finger.Tool != this || finger.Press == null)
            {
                return;
            }

            if (e != null && !finger.PressDelivered)
            {
                // a tap where it could mean more than one thing asks which
                var press = finger.Press;
                IReadOnlyList<object> options = null;
                AsTouch(() => options = SafeFindClickOptions(press));
                if (options.Count > 1)
                {
                    AskWhatTapMeans(sender, press, e, options);
                    return;
                }

                // a tap: the press and the release in one go
                finger.PressDelivered = true;
                AsTouch(() => SafeMouseDown(sender, press));
            }

            var release = e ?? finger.Last;
            if (finger.PressDelivered && release != null)
            {
                AsTouch(() => SafeMouseUp(sender, release));
            }
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
            StartOverMoves();
            currentModifiers = e.KeyModifiers;
            if (e.Pointer.Type == PointerType.Touch)
            {
                TouchReleased(sender, e.Pointer, e);
                return;
            }

            if (rightPress != null && (isSecondaryClickHeld || e.InitialPressMouseButton == MouseButton.Right))
            {
                ReleaseRightButton(sender, e, isReleased: true);
            }
            else if (isSecondaryClickHeld)
            {
                isSecondaryClickHeld = false;
            }
            else if (e.InitialPressMouseButton == MouseButton.Left)
            {
                SafeMouseUp(sender, e);
            }

            // what the press did may have changed what a click here does (a dragged point
            // dropped onto a figure no longer shows the halo of its snap)
            UpdateClickPreview(e);
        }

        void PointerWheelHandler(object sender, PointerWheelEventArgs e)
        {
            HandleWaitingMove();
            currentModifiers = e.KeyModifiers;
            clickPreview.Clear();
            MouseWheel(sender, e);
        }

        public virtual void KeyDown(object sender, KeyEventArgs e)
        {
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
