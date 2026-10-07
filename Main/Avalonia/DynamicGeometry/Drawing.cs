using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using System.Xml.Linq;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    [PropertyGridName("Drawing")]
    public partial class Drawing : IThemeOverridable, IConditionalProperties, INotifyPropertyChanged
    {
        /// <summary>The paper changed: the grid's Background row and its buttons follow</summary>
        public event PropertyChangedEventHandler PropertyChanged;

        public Drawing(Canvas canvas)
        {
            Check.NotNull(canvas, "canvas");

            ActionManager = new ActionManager();
            StyleManager = new StyleManager(this);

            Figures = new RootFigureList(this);

            OnAttachToCanvas += Drawing_OnAttachToCanvas;
            OnDetachFromCanvas += Drawing_OnDetachFromCanvas;

            Canvas = canvas;

            CoordinateSystem = new CoordinateSystem(this);
            // the grid is the drawing's own: a new drawing starts without one, a file says
            CoordinateGrid = new CartesianGrid() { Drawing = this, Visible = false };
            Figures.Add(CoordinateGrid);
            Version = Settings.CurrentDrawingVersion;
        }

        public double Version { get; set; }

        readonly AxisLine[] axisLines = new AxisLine[2];

        /// <summary>
        /// The drawing's x- or y-axis as a line to build on (<see cref="AxisLine"/>): the same
        /// object for the drawing's whole life, in its list or not
        /// </summary>
        public AxisLine GetAxisLine(AxisDirection direction)
        {
            ref var axis = ref axisLines[(int)direction];
            if (axis == null)
            {
                axis = new AxisLine(direction) { Drawing = this };
            }

            return axis;
        }

        /// <summary>The axis lines a click can take that aren't in the list: hit testing looks at them too</summary>
        public IEnumerable<IFigure> UnlistedAxisLines()
        {
            if (CoordinateGrid == null || !CoordinateGrid.ShowsAxes)
            {
                yield break;
            }

            foreach (var direction in new[] { AxisDirection.X, AxisDirection.Y })
            {
                var axis = GetAxisLine(direction);
                if (!Figures.Contains(axis))
                {
                    yield return axis;
                }
            }
        }

        Brush background;

        /// <summary>
        /// The paper: a solid color or a gradient of the drawing's own, or the theme's paper
        /// (null, which files leave out). Under a theme other than the base one, the paper the
        /// drawing chose for that theme, if it did (<see cref="Overrides"/>). Part of the
        /// drawing (saved with it, undoable through the property grid), painted onto whatever
        /// canvas shows it.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridCustomValueProvider(typeof(PaperValue))]
        public Brush Background
        {
            get
            {
                return GetOwnBackground(AppTheme.Current) ?? new SolidColorBrush(AppTheme.Current.Paper);
            }
            set
            {
                background = value;
                ApplyBackground();
            }
        }

        /// <summary>The paper the drawing chose (under the base theme), null for the theme's</summary>
        public Brush OwnBackground
        {
            get
            {
                return background;
            }
        }

        /// <summary>The paper the drawing has under the theme, null for the theme's own</summary>
        public Brush GetOwnBackground(AppTheme theme)
        {
            if (Overrides.TryGetValue(theme.Name, out var values) && values.TryGetValue(nameof(Background), out var overridden))
            {
                return (Brush)overridden;
            }

            return background;
        }

        /// <summary>
        /// By theme name, the paper the drawing chose for that theme (the one property that
        /// can differ, <see cref="Background"/>; null for the theme's paper)
        /// </summary>
        public Dictionary<string, Dictionary<string, object>> Overrides { get; } = new Dictionary<string, Dictionary<string, object>>();

        public void SetOverride(string theme, string property, object value)
        {
            if (!Overrides.TryGetValue(theme, out var values))
            {
                values = new Dictionary<string, object>();
                Overrides[theme] = values;
            }

            values[property] = value;
            ApplyBackground();
        }

        public void RemoveOverride(string theme, string property)
        {
            if (Overrides.TryGetValue(theme, out var values) && values.Remove(property))
            {
                if (values.Count == 0)
                {
                    Overrides.Remove(theme);
                }

                ApplyBackground();
            }
        }

        public void ClearOverrides(string theme)
        {
            Overrides.Remove(theme);
            ApplyBackground();
        }

        /// <summary>
        /// The paper as the property grid sets it: undoing the first choice of a paper gives
        /// back the theme's (no paper of the drawing's own), not the color the theme had
        /// </summary>
        public class PaperValue : PropertyValue, IRestorableValue
        {
            public object CaptureState()
            {
                return ((Drawing)Parent).OwnBackground;
            }

            public void RestoreState(object state)
            {
                ((Drawing)Parent).Background = (Brush)state;
            }
        }

        /// <summary>
        /// The theme's own paper: under the base theme no paper of the drawing's own, under
        /// another an override saying so. An undo step like a paper picked in the grid.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Reset to default")]
        [PropertyGridDestructive]
        [PropertyGridLiveCondition]
        public void UseThemePaper()
        {
            // the theme's already: no undo step that undoes nothing
            if (IsThemePaper)
            {
                return;
            }

            var paper = ThemedValue.ForCurrentTheme(PropertyDiscoveryStrategy.CreateValueProvider(this, nameof(Background)));
            Actions.SetProperty(ActionManager, paper, value: null);
        }

        /// <summary>Whether the paper on screen is the theme's own, not one the drawing chose for the theme</summary>
        bool IsThemePaper
        {
            get
            {
                return !AppTheme.IsBase(AppTheme.Current)
                    && Overrides.TryGetValue(AppTheme.Current.Name, out var values)
                    && values.TryGetValue(nameof(Background), out var chosen)
                    ? chosen == null
                    : OwnBackground == null;
            }
        }

        /// <summary>Drops the paper chosen for the theme on screen: the base theme's again. One undo step.</summary>
        [PropertyGridVisible]
        [PropertyGridIcon(PropertyGridIcon.Cross)]
        [PropertyGridLiveCondition]
        public void SameAsBaseTheme()
        {
            string theme = AppTheme.Current.Name;
            if (!Overrides.TryGetValue(theme, out var values))
            {
                return;
            }

            var saved = new Dictionary<string, object>(values);
            ActionManager.RecordAction(new CallMethodAction(
                () => ClearOverrides(theme),
                () =>
                {
                    foreach (var pair in saved)
                    {
                        SetOverride(theme, pair.Key, pair.Value);
                    }
                }));
        }

        public bool CanEdit(string propertyName)
        {
            if (propertyName == nameof(SameAsBaseTheme))
            {
                var theme = AppTheme.Current;
                return !AppTheme.IsBase(theme) && Overrides.TryGetValue(theme.Name, out var values) && values.Count > 0;
            }

            if (propertyName == nameof(UseThemePaper))
            {
                return !IsThemePaper;
            }

            return true;
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return propertyName == nameof(SameAsBaseTheme) ? "Same paper as in " + AppTheme.Base.Name : defaultCaption;
        }

        /// <summary>
        /// Whether a paper read from a file is plain white: the readers of foreign formats take
        /// that as no paper of the drawing's own, so the theme's shows
        /// </summary>
        public static bool IsWhite(Brush brush)
        {
            return brush is SolidColorBrush solid && solid.Color == Colors.White;
        }

        /// <summary>
        /// Whether the paper is painted onto the canvas at all: a gallery tile shows through
        /// instead, its plate being the drawing's paper or a pastel
        /// </summary>
        public bool PaintsPaper { get; set; } = true;

        /// <summary>
        /// The theme on screen changed, or a color of a theme: the paper, the styles built from
        /// the theme and every figure are drawn again as the theme now says. Every figure
        /// once: the styles' own change notifications, which would have each figure on a style
        /// repaint per property read from the theme, are held back meanwhile
        /// (<see cref="IsRefreshingTheme"/>). Not for a drawing without a canvas (the user's,
        /// parked while they look at the gallery): it catches up when it gets one.
        /// </summary>
        public void RefreshTheme(bool colorsChanged)
        {
            if (Canvas == null)
            {
                return;
            }

            themeVersion = AppTheme.Version;
            IsRefreshingTheme = true;
            try
            {
                if (colorsChanged)
                {
                    StyleManager.RefreshTheme();
                }

                PaintPaper();
            }
            finally
            {
                IsRefreshingTheme = false;
            }

            foreach (var figure in Figures)
            {
                figure.ApplyStyle();
            }
        }

        /// <summary>
        /// While set, figures don't repaint on a change of their style: the drawing applies
        /// every style once at the end (<see cref="RefreshTheme"/>)
        /// </summary>
        public bool IsRefreshingTheme { get; private set; }

        int themeVersion = AppTheme.Version;

        /// <summary>
        /// A drawing that was off screen (a hidden gallery tile, the user's drawing parked while
        /// they looked at the gallery) missed the theme changes meanwhile: refreshes it if there
        /// were any
        /// </summary>
        public void RefreshThemeIfStale()
        {
            if (themeVersion != AppTheme.Version)
            {
                RefreshTheme(colorsChanged: true);
            }
        }

        /// <summary>
        /// Suggested views, in logical coordinates: what the drawing wants to show when it is
        /// opened (a landscape one and maybe a portrait one), for drawings whose content has no
        /// useful bounds - a scene with ground that goes on forever. Whoever fits the drawing
        /// picks the one nearest in shape to the room it has (<see cref="ChooseScene"/>).
        /// </summary>
        public List<Rect> Scenes { get; } = new List<Rect>();

        /// <summary>
        /// Labels that can't be dragged: a drag on one moves the view, as on the paper, and one
        /// on a pinned label (a caption) takes the pinned labels along, so text and figure
        /// scroll together. For the drawings of the gallery, where a thumb on the text of a
        /// phone means scrolling to read the rest. Only the text the drawing came with
        /// (<see cref="FixLabels"/>): a label of a figure (a point's name, a measurement) and a
        /// label the reader adds are theirs to move. Not saved. Not <see cref="IFigure.Locked"/> either: a point counts as
        /// locked when anything built on it is, and a caption with live numbers is built on
        /// its points. The show/hide boxes the drawing came with too: a click ticks one, a
        /// drag on it moves the view.
        /// </summary>
        public HashSet<ControlBase> FixedLabels { get; } = new HashSet<ControlBase>();

        /// <summary>
        /// Makes the text labels the drawing has now <see cref="FixedLabels"/>: the caption and
        /// any text in the plane, not the labels that sit by a figure (a point's name, a
        /// measurement), which a reader moves out of the way of a figure dragged under them;
        /// and its show/hide boxes
        /// </summary>
        public void FixLabels()
        {
            FixedLabels.Clear();
            FixedLabels.UnionWith(Figures.OfType<Label>());
            FixedLabels.UnionWith(Figures.OfType<ShowHideControl>());
        }

        /// <summary>
        /// On while a file is being read: the figures keep the names the file gives them
        /// until all of them are in (expressions are compiled by those names as the figures
        /// come in); <see cref="FigureBase.SettleDefaultNames(Drawing, IFigure)"/> waits. And
        /// the figure list is checked once all of it is in, not as each figure comes in
        /// (<see cref="IFigureExtensions.RecalculateAllDependents"/>).
        /// </summary>
        public bool IsReading { get; set; }

        /// <summary>
        /// On while the figures are worked out again after a move (a drag, a pan by the keys,
        /// undo and redo of either): a label shows its new text at most every
        /// <see cref="LabelBase.TextInterval"/> then (<see cref="LabelBase.ProcessedText"/>)
        /// </summary>
        public bool IsMoving { get; set; }

        Rect? activeScene;

        /// <summary>
        /// The scene that was fitted, if any: a gradient paper spans it rather than the canvas
        /// and is solid beyond it, so the sky doesn't get lighter when the view is zoomed out.
        /// </summary>
        public Rect? ActiveScene
        {
            get
            {
                return activeScene;
            }
            set
            {
                activeScene = value;
                ApplyBackground();
            }
        }

        /// <summary>The scene nearest in shape to a room of this size; null without scenes</summary>
        public Rect? ChooseScene(double roomWidth, double roomHeight)
        {
            if (Scenes.Count == 0 || roomWidth <= 0 || roomHeight <= 0)
            {
                return null;
            }

            double room = System.Math.Log(roomWidth / roomHeight);
            return Scenes.OrderBy(scene => System.Math.Abs(System.Math.Log(scene.Width / scene.Height) - room)).First();
        }

        /// <summary>Fits the scene into the canvas edge to edge and makes it the active one</summary>
        public void ShowScene(Rect scene)
        {
            ActiveScene = scene;
            CoordinateSystem.FitScene(scene);
        }

        /// <summary>
        /// "Zoom to fit" as whoever shows the drawing does it, when that takes more than the
        /// bounds of the content: where a caption leaves room for the figure, which part of
        /// the plane a graph is about. Null for the plain fit.
        /// </summary>
        public Action FitToWindow { get; set; }

        /// <summary>
        /// Zoom to fit, as the user asks for it (a double click, H, the context menu): the
        /// layout the drawing was opened with, a scene if it has scenes, else everything
        /// visible. (It was the last of these always: a double click on a drawing of the
        /// gallery put the figure under its caption, zoomed a graph onto its few points
        /// and showed the Castle's points instead of its scene.)
        /// </summary>
        public void ZoomToFit()
        {
            if (FitToWindow != null)
            {
                FitToWindow();
                return;
            }

            var scene = Canvas != null ? ChooseScene(Canvas.Bounds.Width, Canvas.Bounds.Height) : null;
            if (scene != null)
            {
                ShowScene(scene.Value);
            }
            else
            {
                CoordinateSystem.ZoomExtend();
            }
        }

        void ApplyBackground()
        {
            PaintPaper();
            CoordinateGrid?.ApplyStyle();
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Background)));
        }

        /// <summary>The paper onto the canvas, and the grid's colors, which go by the paper it is on</summary>
        void PaintPaper()
        {
            CoordinateGrid?.RefreshTheme();
            if (Canvas != null)
            {
                Canvas.Background = PaintsPaper ? PlaceBackground(Background) : null;
            }
        }

        /// <summary>
        /// A gradient's relative start and end are relative to the active scene, not to the
        /// canvas: the brush the canvas gets has them in pixels, and pads with the end colors.
        /// </summary>
        Brush PlaceBackground(Brush brush)
        {
            if (activeScene == null || CoordinateSystem == null || !(brush is LinearGradientBrush gradient))
            {
                return brush;
            }

            var scene = activeScene.Value;
            var topLeft = CoordinateSystem.ToPhysical(new Point(scene.X, scene.Bottom)); // y up: Bottom is the top edge
            var bottomRight = CoordinateSystem.ToPhysical(new Point(scene.Right, scene.Y));
            var placed = new LinearGradientBrush()
            {
                StartPoint = new RelativePoint(Place(gradient.StartPoint), RelativeUnit.Absolute),
                EndPoint = new RelativePoint(Place(gradient.EndPoint), RelativeUnit.Absolute),
                SpreadMethod = GradientSpreadMethod.Pad
            };
            foreach (var stop in gradient.GradientStops)
            {
                placed.GradientStops.Add(new GradientStop(stop.Color, stop.Offset));
            }

            return placed;

            Point Place(RelativePoint point)
            {
                if (point.Unit == RelativeUnit.Absolute)
                {
                    return point.Point;
                }

                return new Point(
                    topLeft.X + point.Point.X * (bottomRight.X - topLeft.X),
                    topLeft.Y + point.Point.Y * (bottomRight.Y - topLeft.Y));
            }
        }

        void Drawing_OnAttachToCanvas(Canvas canvas)
        {
            canvas.Background = PaintsPaper ? PlaceBackground(Background) : null;
            canvas.SizeChanged += mCanvas_SizeChanged;
            UpdateClip(canvas);
            foreach (var figure in Figures)
            {
                figure.OnAddingToCanvas(canvas);
            }

            // parked while the theme changed (a switch, a tweaked color): the shapes still wear the old look
            RefreshThemeIfStale();
        }

        void UpdateClip(Canvas canvas)
        {
            canvas.Clip = new RectangleGeometry() { Rect = new Rect(0, 0, canvas.ActualWidth, canvas.ActualHeight) };
        }

        void Drawing_OnDetachFromCanvas(Canvas canvas)
        {
            canvas.SizeChanged -= mCanvas_SizeChanged;
            foreach (var figure in Figures)
            {
                figure.OnRemovingFromCanvas(canvas);
            }
            this.Behavior = null;
        }

        #region Events

        public event EventHandler<SelectionChangedEventArgs> SelectionChanged;

        public class SelectionChangedEventArgs : EventArgs
        {
            public SelectionChangedEventArgs()
            {
                SelectedFigures = Enumerable.Empty<IFigure>();
            }

            public SelectionChangedEventArgs(IEnumerable<IFigure> selection)
                : this()
            {
                SelectedFigures = selection;
            }

            public SelectionChangedEventArgs(IFigure singleSelection)
                : this(singleSelection.AsEnumerable())
            {
            }

            public IEnumerable<IFigure> SelectedFigures { get; set; }
        }

        public event EventHandler<DeleteExecutedEventArgs> DeleteExecuted;

        public class DeleteExecutedEventArgs : EventArgs
        {
            public DeleteExecutedEventArgs()
            {
                DeletedFigures = Enumerable.Empty<IFigure>();
            }

            public DeleteExecutedEventArgs(IEnumerable<IFigure> deletedFigures)
                : this()
            {
                DeletedFigures = deletedFigures;
            }

            public DeleteExecutedEventArgs(IFigure deletedFigure)
                : this(deletedFigure.AsEnumerable())
            {
            }

            public IEnumerable<IFigure> DeletedFigures { get; set; }
        }

        public void ClearLockedFigures()
        {
            foreach (IFigure figure in this.Figures)
            {
                if (figure.Locked)
                {
                    figure.Locked = false;
                }
            }
        }

        public void RaiseSelectionChanged(SelectionChangedEventArgs args)
        {
            if (SelectionChanged != null)
            {
                SelectionChanged(this, args);
            }
        }

        public void RaiseSelectionChanged(params IFigure[] selected)
        {
            if (SelectionChanged != null)
            {
                SelectionChanged(this, new SelectionChangedEventArgs(selected));
            }
        }

        public void RaiseSelectionChanged(IEnumerable<IFigure> selected)
        {
            if (SelectionChanged != null)
            {
                SelectionChanged(this, new SelectionChangedEventArgs(selected));
            }
        }

        /// <summary>
        /// A tool that picks figures (<see cref="IFigurePicker"/>) picked one or let go of one:
        /// the picks are selected figures, which the Figure List paints again. Not
        /// <see cref="SelectionChanged"/>, which would put the selection's properties in the
        /// side panel over the tool's own.
        /// </summary>
        public event EventHandler PicksChanged;

        public void RaisePicksChanged()
        {
            PicksChanged?.Invoke(this, EventArgs.Empty);
        }

        public void RaiseDeleteExecuted(DeleteExecutedEventArgs args)
        {
            if (DeleteExecuted != null)
            {
                DeleteExecuted(this, args);
            }
        }

        public void RaiseDeleteExecuted(params IFigure[] selected)
        {
            if (DeleteExecuted != null)
            {
                DeleteExecuted(this, new DeleteExecutedEventArgs(selected));
            }
        }

        public void RaiseDeleteExecuted(IEnumerable<IFigure> selected)
        {
            if (DeleteExecuted != null)
            {
                DeleteExecuted(this, new DeleteExecutedEventArgs(selected));
            }
        }

        public class DisplayPropertiesEventArgs : EventArgs
        {
            public object Object { get; set; }

            /// <summary>The property whose editor takes the keyboard, this once; null for none</summary>
            public string FocusProperty { get; set; }
        }

        public event EventHandler<DisplayPropertiesEventArgs> DisplayProperties;

        /// <param name="focusProperty">The property to type into right away (a new label's text)</param>
        public void RaiseDisplayProperties(object objectWithProperties, string focusProperty = null)
        {
            if (DisplayProperties != null)
            {
                DisplayProperties(this, new DisplayPropertiesEventArgs() { Object = objectWithProperties, FocusProperty = focusProperty });
            }
        }

        public class ConstructionStepCompleteEventArgs : EventArgs
        {
            public bool ConstructionComplete { get; set; }
            public bool ConstructionRollback { get; set; }
            public Type FigureTypeNeeded { get; set; }
            public IEnumerable<IFigure> FigureResults { get; set; }
        }

        public class ConstructionStepStartedEventArgs : EventArgs
        {
        }

        public event EventHandler<ConstructionStepStartedEventArgs> ConstructionStepStarted;
        public event EventHandler<ConstructionStepCompleteEventArgs> ConstructionStepComplete;

        public void RaiseConstructionStepComplete(ConstructionStepCompleteEventArgs args)
        {
            if (ConstructionStepComplete != null)
            {
                ConstructionStepComplete(this, args);
            }
        }

        public void RaiseConstructionStepStarted(ConstructionStepStartedEventArgs args)
        {
            if (ConstructionStepStarted != null)
            {
                ConstructionStepStarted(this, args);
            }
        }

        public void RaiseConstructionStepStarted()
        {
            if (ConstructionStepStarted != null)
            {
                ConstructionStepStarted(this, new ConstructionStepStartedEventArgs());
            }
        }

        public class ConstructionFeedbackEventArgs : EventArgs
        {
            public Type FigureTypeNeeded { get; set; }
            public bool IsMouseButtonDown { get; set; }
        }

        public event EventHandler<ConstructionFeedbackEventArgs> ConstructionFeedback;
        public void RaiseConstructionFeedback(ConstructionFeedbackEventArgs args)
        {
            if (ConstructionFeedback != null)
            {
                ConstructionFeedback(this, args);
            }
        }

        // UserIsAddingFigures is intended to occur after figures have been added to drawing but before transaction is committed.
        public event EventHandler<UIAFEventArgs> UserIsAddingFigures;

        public class UIAFEventArgs : EventArgs
        {
            public IEnumerable<IFigure> Figures { get; set; }
        }

        /// <summary>
        /// This offers an opportunity to do additional processing of figures after they have been added to the drawing.
        /// Unlike FigureList.OnItemAdded(), this event does not occur on undo or redo.
        /// </summary>
        public void RaiseUserIsAddingFigures(UIAFEventArgs figures)
        {
            if (UserIsAddingFigures != null)
            {
                UserIsAddingFigures(this, figures);
            }
        }

        public class DocumentOpenRequestedEventArgs : EventArgs
        {
            public enum InWhichWindowChoice
            {
                DontCare,
                ReuseCurrent,
                NewWindowOrTab
            }

            public string DocumentXml { get; set; }
            public InWhichWindowChoice InWhichWindow { get; set; }
        }

        /// <summary>
        /// This event is raised when the user clicks on a hyperlink
        /// to open another drawing document, much like a web-browser link.
        /// This event signals to the host of the drawing to either open this 
        /// new drawing in a separate tab or replace the current one.
        /// </summary>
        public event EventHandler<DocumentOpenRequestedEventArgs> DocumentOpenRequested;
        public void RaiseDocumentOpenRequested(DocumentOpenRequestedEventArgs args)
        {
            if (DocumentOpenRequested != null)
            {
                DocumentOpenRequested(this, args);
            }
        }

        public event SizeChangedEventHandler SizeChanged;
        public void RaiseSizeChanged(SizeChangedEventArgs args)
        {
            if (SizeChanged != null)
            {
                SizeChanged(this, args);
            }
        }

        public event Action<string> Status;
        public void RaiseStatusNotification(string status)
        {
            if (Status != null)
            {
                Status(status);
            }
        }

        /// <summary>
        /// What a click at the cursor would take, while it could take more than one thing
        /// (<see cref="ClickChoice"/>): shown over the status, which comes back with null
        /// </summary>
        public event Action<string> ChoiceStatus;
        public void RaiseChoiceStatus(string text)
        {
            ChoiceStatus?.Invoke(text);
        }

        public event Action ZoomChanged;    // Used by Tabula.
        public void RaiseZoomChanged()
        {
            if (ZoomChanged != null)
            {
                ZoomChanged();
            }
        }

        public event EventHandler<FigureCoordinatesChangedEventArgs> FigureCoordinatesChanged;
        public class FigureCoordinatesChangedEventArgs : EventArgs
        {
            public FigureCoordinatesChangedEventArgs()
            {
                Figures = Enumerable.Empty<IFigure>();
            }

            public FigureCoordinatesChangedEventArgs(IEnumerable<IFigure> figures)
                : this()
            {
                Figures = figures;
            }

            public FigureCoordinatesChangedEventArgs(IFigure singleFigure)
                : this(singleFigure.AsEnumerable())
            {
            }

            public IEnumerable<IFigure> Figures { get; set; }
        }

        public void RaiseFigureCoordinatesChanged(FigureCoordinatesChangedEventArgs args)
        {
            if (FigureCoordinatesChanged != null)
            {
                FigureCoordinatesChanged(this, args);
            }
        }

        #endregion

        public string Name { get; set; }
        public override string ToString()
        {
            return Name;
        }

        /// <summary>
        /// Compares the last item in ActionManager.EnumUndoableActions() to a value stored at the last Save to determine if there are unsaved changes.
        /// </summary>
        public bool HasUnsavedChanges
        {
            get
            {
                var undoableActions = ActionManager.EnumUndoableActions();
                if (undoableActions.IsEmpty())
                {
                    return false;
                }

                if (LastUndoableActionAtSave == undoableActions.Last())
                {
                    return false;
                }
                else
                {
                    return true;
                }

            }
        }

        public IAction LastUndoableActionAtSave { get; set; }
        public ActionManager ActionManager { get; set; }

        /// <summary>
        /// A tool is in the middle of a construction, or a point is being dragged: whatever is
        /// recorded now joins that undo step, and is taken back with it if it is given up.
        /// Commands that are not part of it (Delete, Paste, Redo) wait.
        /// </summary>
        public bool IsRecordingTransaction
        {
            get { return ActionManager.RecordingTransaction != null; }
        }

        public event Action<Canvas> OnAttachToCanvas;
        public event Action<Canvas> OnDetachFromCanvas;

        public event Action<Behavior> BehaviorChanged;

        /// <summary>
        /// Informs the clients that the current behavior of the drawing
        /// was set to a new behavior.
        /// </summary>
        /// <param name="behavior">The new behavior of the drawing.</param>
        void RaiseBehaviorChanged(Behavior behavior)
        {
            if (BehaviorChanged != null)
            {
                BehaviorChanged(behavior);
            }
        }

        private Canvas mCanvas;
        public Canvas Canvas
        {
            get
            {
                return mCanvas;
            }
            set
            {
                if (mCanvas == value)
                {
                    return;
                }
                if (mCanvas != null && OnDetachFromCanvas != null)
                {
                    OnDetachFromCanvas(mCanvas);
                }
                mCanvas = value;
                if (mCanvas != null && OnAttachToCanvas != null)
                {
                    OnAttachToCanvas(mCanvas);
                }
            }
        }

        void mCanvas_SizeChanged(object sender, Avalonia.Controls.SizeChangedEventArgs e)
        {
            UpdateClip(Canvas);
            RaiseSizeChanged(e);
        }

        public StyleManager StyleManager { get; set; }

        private Behavior mBehavior;
        public Behavior Behavior
        {
            get
            {
                return mBehavior;
            }
            set
            {
                if (mBehavior == value)
                {
                    return;
                }
                if (mBehavior != null)
                {
                    mBehavior.Stopping();
                    mBehavior.Drawing = null;
                }
                mBehavior = value;
                if (mBehavior != null)
                {
                    mBehavior.Drawing = this;
                    mBehavior.Started();
                    RaiseBehaviorChanged(mBehavior);
                }
                
            }
        }

        public void SetDefaultBehavior()
        {
            Behavior = Behavior.Default;
        }

        public CoordinateSystem CoordinateSystem { get; set; }
        public CartesianGrid CoordinateGrid { get; set; }

        public RootFigureList Figures { get; set; }

        public IEnumerable<IFigure> GetSerializableFigures()
        {
            foreach (IFigure figure in Figures.Where(f=>f.Serializable))
            {
                yield return figure;
            }
        }

        /// <summary>
        /// The selection as the grid shows it: the figures selected, and the vertices and sides
        /// of a regular polygon selected by themselves (<see cref="FigureParts"/>). Whatever
        /// acts on figures of the drawing takes <see cref="FigureParts.Wholes"/> of it.
        /// </summary>
        public IEnumerable<IFigure> GetSelectedFigures()
        {
            foreach (IFigure figure in Figures)
            {
                if (figure.Selected)
                {
                    yield return figure;
                }
                else if (figure is IFigureParts parts)
                {
                    foreach (var part in parts.SelectableParts.Where(part => part.Selected))
                    {
                        yield return part;
                    }
                }
            }
        }

        public Point GetSelectionCenter()
        {
            Point center = new Point(0, 0);
            var selectedFigures = GetSelectedFigures();
            var count = selectedFigures.Count();
            if (count > 0)
            {
                foreach (var figure in selectedFigures)
                {
                    center += new Avalonia.Vector(figure.Center.X, figure.Center.Y);
                }
                center = new Point(center.X / count, center.Y / count);
            }
            return center;
        }

        public List<IFigure> GetSelectedFiguresWithDependencies()
        {
            var selectedFigures = FigureParts.Wholes(GetSelectedFigures()).ToList();
            List<IFigure> results = new List<IFigure>();
            results.AddRange(selectedFigures);
            foreach (IFigure selectedFigure in selectedFigures)
            {
                foreach (IFigure f in Figures)
                {
                    if (selectedFigure.DependsOn(f) && !results.Contains(f))
                    {
                        results.Add(f);
                    }
                }
            }
            return results;
        }

        public IEnumerable<IFigure> GetLockedFigures()
        {
            foreach (IFigure figure in Figures)
            {
                if (figure.Locked)
                {
                    yield return figure;
                }
            }
        }

#if !SILVERLIGHT
        public static Drawing Load(string path, Canvas canvas)
        {
            Drawing drawing = new Drawing(canvas);
            new DrawingDeserializer().ReadDrawing(drawing, System.IO.File.ReadAllText(path));
            return drawing;
        }

        public void Save(string path)
        {
            DrawingSerializer.Save(this, path);
        }
#endif

        [Obsolete("Use Actions.Add instead")]
        public void Add(IFigure figure)
        {
            Actions.Add(this, figure);
        }

#if !PLAYER

        [Obsolete("Use Actions.Remove instead")]
        public void Remove(IFigure figure)
        {
            Actions.Remove(figure);
        }

#endif

        public void Recalculate()
        {
            // the paper is pinned to the scene, so it moves with the view
            if (activeScene != null)
            {
                ApplyBackground();
            }

            foreach (var figure in Figures)
            {
                figure.RecalculateAndUpdateVisual();
            }
        }

        [Obsolete("Use Actions.Add instead")]
        public void Add(IEnumerable<IFigure> figures)
        {
            using (Transaction.Create(ActionManager))
            {
                foreach (var figure in figures)
                {
                    Actions.Add(this, figure);
                }
            }
        }

        /// <summary>
        /// What of the file the drawing was read from could not be read, a line each; null
        /// when all of it was. Such a drawing is not the file any more: it does not take the
        /// file's name, and Save does not write what is left of it over the file.
        /// </summary>
        public string LoadErrors { get; private set; }

        public void AddFromXml(XElement element)
        {
            var deserializer = new DrawingDeserializer();
            deserializer.ReadDrawing(this, element);
            LoadErrors = deserializer.IsSuccess ? null : deserializer.GetErrorReport();
            if (LoadErrors != null)
            {
                var lines = LoadErrors.Split('\n');
                RaiseStatusNotification(
                    "Not all of this file could be read. "
                    + lines[0]
                    + (lines.Length > 1 ? " (and " + (lines.Length - 1) + " more)" : ""));
            }
        }

#if !TABULAPLAYER
        public void AddFromDGF(string[] lines)
        {
            var reader = new DGFReader();
            reader.ReadDrawing(this, lines);
            if (!reader.IsSuccess)
            {
                RaiseStatusNotification(reader.GetErrorReport());
            }
        }

        /// <summary>
        /// A GeoGebra worksheet (the geogebra.xml of a .ggb). That something couldn't be read
        /// is said in the status; what exactly goes to the console, the list can be long.
        /// </summary>
        public void AddFromGeoGebra(XElement worksheet)
        {
            var reader = new GeoGebraReader();
            reader.ReadDrawing(this, worksheet);
            if (!reader.IsSuccess)
            {
                Console.WriteLine("GeoGebra: " + reader.Details.Replace(Environment.NewLine, Environment.NewLine + "GeoGebra: "));
                RaiseStatusNotification(reader.GetErrorReport());
            }
        }
#endif
#if !PLAYER

        public string SaveAsText()
        {
            return DrawingSerializer.SaveDrawing(this);
        }

        /// <summary>
        /// Each selected figure is removed the way the property grid's Delete button removes
        /// one (<see cref="RemoveFigureAction"/>: dependents go with it, a polygon loses the
        /// vertex instead of dying, point labels come back on undo, an auxiliary Number goes
        /// with its last user), all in one undo step.
        /// </summary>
        public void DeleteSelection()
        {
            Delete(this.GetSelectedFigures());
        }

        /// <summary>Several figures in one undo step, as <see cref="DeleteSelection"/> does</summary>
        public void Delete(IEnumerable<IFigure> figuresToDelete)
        {
            // a vertex or a side goes with its polygon: a regular polygon without one is none
            var figures = FigureParts.Wholes(figuresToDelete)
                .Where(f => !(f is CartesianGrid) && !(f is PointLabel) && !(f is FigureLabel))
                .ToArray();
            // the Delete key reaches here twice (the tool on key down, the window on key up):
            // the second time there is nothing selected
            if (figures.Length == 0)
            {
                return;
            }

            // a tool in the middle of a construction has its transaction open: a deletion now
            // would be part of that figure's undo step, and Escape would put the figures back
            if (IsRecordingTransaction)
            {
                return;
            }

            // not delayed: each removal happens now, so that a figure that went as a
            // dependent of an earlier one is seen to be gone and not removed a second time
            using (Transaction.Create(ActionManager, false))
            {
                foreach (var figure in figures)
                {
                    if (Figures.Contains(figure))
                    {
                        Actions.Remove(figure);
                    }
                }
            }
        }

#endif

#if !SILVERLIGHT

        /// <summary>In pixels: how far down and to the right each paste puts its copy from the one before</summary>
        public const double PasteStep = 24;

        // the pastes of what is on the clipboard so far: each goes a step further, or the
        // second paste would lie exactly on the first
        static int pastes;

        public void Copy()
        {
            List<IFigure> list = new List<IFigure>(this.GetSelectedFiguresWithDependencies());

            // nothing selected: the clipboard keeps what it has (it was emptied)
            if (list.Count == 0)
            {
                return;
            }

            pastes = 0;
            var s = new System.Text.StringBuilder();
            using (var w = System.Xml.XmlWriter.Create(s, new System.Xml.XmlWriterSettings()
            {
                Indent = true
            }))
            {
                new DrawingSerializer().WriteFiguresWithStyles(this, list, w);
            }
            Clipboard.SetText(s.ToString());
        }

        /// <summary>
        /// Whether the text can be figures: what Copy puts on the clipboard, or a drawing's
        /// file. Text copied from anywhere else is not, and reading it as figures threw - an
        /// error report for a Ctrl+V. (A check, not a caught exception: every exception is
        /// shown.)
        /// </summary>
        static bool CanBeFigures(string text)
        {
            var start = text.TrimStart();
            if (start.StartsWith("<?xml", StringComparison.Ordinal))
            {
                int end = start.IndexOf("?>", StringComparison.Ordinal);
                start = end < 0 ? "" : start.Substring(end + 2).TrimStart();
            }

            return start.StartsWith("<Figures", StringComparison.Ordinal) || start.StartsWith("<Drawing", StringComparison.Ordinal);
        }

        public void Paste()
        {
            if (Clipboard.GetText() != null)
            {
                this.PasteFromText(Clipboard.GetText());
            }
        }

        public void PasteFromText(string str)
        {
            // Not into the undo step of a construction under way. A step down and to the
            // right of the originals, selected: on top of them and not selected, a paste
            // looked like nothing at all, and each try left another hidden copy.
            if (str != null && !IsRecordingTransaction)
            {
                if (!CanBeFigures(str))
                {
                    RaiseStatusNotification("Nothing to paste: copy figures first (select them, then Ctrl+C).");
                    return;
                }

                pastes++;
                Actions.Paste(this, str, pixelOffset: pastes * PasteStep);
            }
        }

        public void PasteFrom(string xmlFile)
        {
            string copiedFigures = System.IO.File.ReadAllText(xmlFile);
            PasteFromText(copiedFigures);
        }

#endif
#if !PLAYER
        public void Duplicate()
        {
            var transaction = Transaction.Create(ActionManager, false);
            List<IFigure> list = GetSelectedFiguresWithDependencies();
            var s = new System.Text.StringBuilder();
            using (var w = System.Xml.XmlWriter.Create(s, new System.Xml.XmlWriterSettings()
            {
                Indent = true
            }))
            {
                new DrawingSerializer().WriteFigureList(list, w);
            }
            var paste = new PasteAction(this, s.ToString());
            ActionManager.RecordAction(paste);

            // Offset the new figures.
            List<IMovable> moving = new List<IMovable>();
            foreach (IFigure f in paste.Figures.Where(f => f as IMovable != null))
            {
                moving.Add(f as IMovable);
            }
            Actions.Move(this, moving, new Point(1, -1), Figures);

            // Select new figures.
            Figures.ClearSelection();
            foreach (IFigure f in paste.Figures)
            {
                f.Selected = true;
            }

            RaiseUserIsAddingFigures(new UIAFEventArgs() {Figures = paste.Figures});
            transaction.Commit();
        }

#endif
      
        public void SelectAll()
        {
            foreach (IFigure figure in Figures)
            {
                // nor a Number, which has nothing on the paper: a selection shows only the
                // rows all its figures have, and a number has no Visible or Locked
                if (!(figure is CartesianGrid) && !(figure is Number))
                {
                    figure.Selected = true;
                }
            }
        }

        public void LockSelected()
        {
            IEnumerable<IFigure> roots = FigureParts.Wholes(this.GetSelectedFigures()).ToList();
            bool shouldLock = roots.All(root => (!root.Locked));
            foreach (IFigure figure in roots)
            {
                if (!(figure is CartesianGrid))
                {
                    figure.Locked = shouldLock;
                }
            }
        }

        public void ClearStatus()
        {
            RaiseStatusNotification("");
        }

        public event EventHandler<UnhandledExceptionNotificationEventArgs> UnhandledException;
        public void RaiseError(object sender, Exception ex)
        {
            if (UnhandledException != null)
            {
                UnhandledException(sender, new UnhandledExceptionNotificationEventArgs(ex));
            }
        }
    }

    public class UnhandledExceptionNotificationEventArgs : EventArgs
    {
        public UnhandledExceptionNotificationEventArgs(Exception ex)
        {
            Exception = ex;
        }

        public Exception Exception { get; set; }
    }
}
