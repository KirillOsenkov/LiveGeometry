using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry
{
    public class DrawingHost : Grid
    {
        public event EventHandler ReadyForInteraction;
        public event EventHandler<UnhandledExceptionNotificationEventArgs> UnhandledException = delegate { };

        public Drawing CurrentDrawing
        {
            get
            {
                return this.DrawingControl.Drawing;
            }
        }

        public Ribbon Ribbon { get; set; }
        public DrawingControl DrawingControl { get; set; }
        public PropertyGrid PropertyGrid { get; set; }
        public StatusBar StatusBar { get; set; }
        public FigureExplorer FigureExplorer { get; set; }

        protected ScrollViewer propertyGridScrollViewer;

        /// <summary>The side panel: the title of the property grid and a close cross, the rows scrolling under them</summary>
        Border sidePanel;

        public Command CommandToggleGrid { get; set; }
        public Command CommandToggleOrtho { get; set; }
        public Command CommandToggleSnapToGrid { get; set; }
        public Command CommandToggleSnapToPoint { get; set; }
        public Command CommandToggleLabelNewPoints { get; set; }
        public Command CommandTogglePolar { get; set; }
        public Command CommandToggleSnapToCenter { get; set; }
        public Command CommandTogglePointByCoordinates { get; set; }
        public Command CommandToggleFigureExplorer { get; set; }

        GridSplitter figureExplorerSplitter;

        /// <summary>The width the Figure List had when it was put away, to come back with</summary>
        double figureExplorerWidth = 240;

        public DrawingHost()
        {
            Behavior.NewBehaviorCreated += Behavior_NewBehaviorCreated;
            Behavior.BehaviorDeleted += Behavior_BehaviorDeleted;
            SetupLayout();

            // the side panel is built for the theme it opened under: the style swatches are
            // drawn for it, and an edit goes to that theme's values (ThemedValue). The colors
            // of a theme are those of the theme on screen: the page goes over to the new one.
            AppTheme.CurrentChanged += () =>
            {
                var selection = PropertyGrid.Selection;
                if (selection is AppTheme)
                {
                    selection = AppTheme.Current;
                }

                if (selection != null && CurrentDrawing != null)
                {
                    ShowProperties(selection);
                }
            };
        }

        protected virtual void SetupLayout()
        {
            // the Figure List | its splitter | the canvas, with the side panel and the status bar over the canvas
            this.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
            this.RowDefinitions.Add(new RowDefinition());
            this.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(0), MaxWidth = 600 });
            this.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
            this.ColumnDefinitions.Add(new ColumnDefinition());

            CreateRibbon();
            CreateCanvas();
            CreatePropertyGrid();
            CreateStatusBar();
            CreateFigureExplorer();

            this.Children.Add(Ribbon);
            this.Children.Add(FigureExplorer);
            this.Children.Add(figureExplorerSplitter);
            this.Children.Add(DrawingControl);
            this.Children.Add(sidePanel);
            this.Children.Add(StatusBar);

            Grid.SetColumnSpan(Ribbon, 3);
            Grid.SetRow(FigureExplorer, 1);
            Grid.SetRow(figureExplorerSplitter, 1);
            Grid.SetColumn(figureExplorerSplitter, 1);
            foreach (var control in new Control[] { DrawingControl, sidePanel, StatusBar })
            {
                Grid.SetRow(control, 1);
                Grid.SetColumn(control, 2);
            }

            ShowFigureExplorer(Settings.Instance.ShowFigureExplorer);

            CommandToggleGrid = new Command(ToggleGrid, CartesianGrid.GetIcon(), "Grid", BehaviorCategories.Coordinates)
            {
                IsChecked = () => CurrentDrawing != null && CurrentDrawing.CoordinateGrid.Visible,
                Shortcut = "G",
                HintText = "Show coordinate axes and grid. "
                    + "Hold Shift while placing or dragging a point to snap it to the grid, with or without the grid on screen."
            };
            CommandToggleOrtho = new Command(ToggleOrtho, ToggleIcons.Ortho(), "Ortho", BehaviorCategories.Selection)
            {
                IsChecked = () => Settings.Instance.EnableOrtho,
                HintText = "The next point of a figure goes exactly to the side of, or straight above or below, the one before it."
            };
            CommandToggleSnapToGrid = new Command(ToggleSnapToGrid, ToggleIcons.SnapToGrid(), "Snap to grid", BehaviorCategories.Selection)
            {
                IsChecked = () => Settings.Instance.EnableSnapToGrid,
                HintText = "New and dragged points land on the grid. Shift does the same for one point."
            };
            CommandToggleSnapToPoint = new Command(ToggleSnapToPoint, ToggleIcons.SnapToPoint(), "Snap to point", BehaviorCategories.Selection)
            {
                IsChecked = () => Settings.Instance.EnableSnapToPoint,
                HintText = "A click near an existing point picks that point instead of making a new one."
            };
            CommandToggleLabelNewPoints = new Command(ToggleLabelNewPoints, ToggleIcons.LabelNewPoints(), "Label new points", BehaviorCategories.Points)
            {
                IsChecked = () => Settings.Instance.AutoLabelPoints,
                HintText = "Every new point shows its name (A, B, C...) beside it. "
                    + "To enable the name later select the point and turn on Show name."
            };
            CommandTogglePolar = new Command(TogglePolar, ToggleIcons.Polar(), "Polar", BehaviorCategories.Selection)
            {
                IsChecked = () => Settings.Instance.EnablePolar,
                HintText = "The next point of a figure goes at a round angle from the one before it."
            };
            CommandToggleSnapToCenter = new Command(ToggleSnapToCenter, ToggleIcons.SnapToCenter(), "Snap to center", BehaviorCategories.Selection)
            {
                IsChecked = () => Settings.Instance.EnableSnapToCenter,
                HintText = "A click near the middle of a segment makes its midpoint."
            };
            CommandTogglePointByCoordinates = new Command(TogglePointByCoordinates, ToggleIcons.PointByCoordinates(), "Point by coordinates", BehaviorCategories.Coordinates)
            {
                IsChecked = () => Settings.Instance.EnablePointByCoordinates,
                HintText = "Enables typing X and Y when a point is needed."
                    + " For a point by itself, "
                    + "the Coordinates tool on the Points tab needs no toggle."
            };
            CommandToggleFigureExplorer = new Command(ToggleFigureExplorer, ToggleIcons.FigureList(), "Figure List", BehaviorCategories.Selection)
            {
                IsChecked = () => FigureExplorer.IsVisible,
                HintText = "A list of all figures in the drawing at the left, hidden ones faded. "
                    + "Selecting a row selects the figure; arrows in the margin show what it is built on."
            };
        }

        /// <summary>The drawing's own properties (its paper) in the side panel</summary>
        public void ShowDrawingProperties()
        {
            if (CurrentDrawing == null)
            {
                return;
            }

            ShowProperties(CurrentDrawing);
        }

        protected void CreateFigureExplorer()
        {
            FigureExplorer = new FigureExplorer();
            figureExplorerSplitter = new GridSplitter()
            {
                Width = 4,
                ResizeDirection = GridResizeDirection.Columns
            };
            figureExplorerSplitter.BindTheme(GridSplitter.BackgroundProperty, nameof(AppTheme.HeaderRow));
        }

        protected virtual void CreateStatusBar()
        {
            StatusBar = new StatusBar();
            StatusBar.HorizontalAlignment = HorizontalAlignment.Left;
            StatusBar.VerticalAlignment = VerticalAlignment.Bottom;
            StatusBar.ZIndex = (int)ZOrder.StatusBar;
        }

        protected void CreatePropertyGrid()
        {
            // the title and the close cross stay put above the rows, which scroll under them
            var title = new Decorator();
            PropertyGrid = new PropertyGrid() { HeaderHost = title };
            PropertyGrid.VisibilityChanged += PropertyGrid_VisibilityChanged;

            var header = new DockPanel() { Margin = new Thickness(14, 10, 8, 0) };
            var closeButton = CreateCloseButton();
            DockPanel.SetDock(closeButton, Dock.Right);
            header.Children.Add(closeButton);
            header.Children.Add(title);

            var rows = new Border()
            {
                Padding = new Thickness(14, 0, 14, 12),
                Child = PropertyGrid
            };

            // A list (the styles, the emoji) brings its selected item into view when it is
            // laid out, and the request goes on up to the panel's own scroll viewer: the
            // panel opened scrolled down to the figure's style, its first rows (a
            // segment's marks) cut off. The list scrolls itself; the panel stays put.
            rows.AddHandler(
                Control.RequestBringIntoViewEvent,
                (sender, e) =>
                {
                    if (e.TargetObject is ListBoxItem)
                    {
                        e.Handled = true;
                    }
                });

            propertyGridScrollViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                Content = rows
            };

            var layout = new DockPanel();
            DockPanel.SetDock(header, Dock.Top);
            layout.Children.Add(header);
            layout.Children.Add(propertyGridScrollViewer);

            // the same surface as the tools strip of the ribbon
            sidePanel = new Border()
            {
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                ClipToBounds = true, // the rows scroll under the rounded corners
                MinWidth = 200.0,
                Margin = new Thickness(8),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Visibility = Visibility.Collapsed,
                ZIndex = (int)ZOrder.StatusBar,
                Child = layout
            };
            sidePanel.BindTheme(Border.BackgroundProperty, nameof(AppTheme.Background));
            sidePanel.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.TabLine));

            PropertyGrid.ValueDiscoveryStrategy = new ExcludeByDefaultValueDiscoveryStrategy();
        }

        // raised whenever the grid is given something to show, by whoever (a figure's
        // "Edit this style" shows the style itself)
        private void PropertyGrid_VisibilityChanged(object sender, EventArgs e)
        {
            sidePanel.Visibility = PropertyGrid.Visibility;

            // a page starts at its top, not where the one before it was scrolled to
            propertyGridScrollViewer.Offset = default;
        }

        /// <summary>A small faint cross that puts the side panel away (<see cref="CloseSidePanel"/>)</summary>
        Control CreateCloseButton()
        {
            var cross = new Avalonia.Controls.Shapes.Path()
            {
                Data = Geometry.Parse("M0,0 L8,8 M8,0 L0,8"),
                StrokeThickness = 1.5,
                StrokeLineCap = PenLineCap.Round,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            var button = new Border()
            {
                Width = 20,
                Height = 20,
                CornerRadius = new CornerRadius(4),
                Background = Brushes.Transparent, // hit-testable around the cross too
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(8, -2, 0, 0),
                Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                Child = cross
            };
            ToolTip.SetTip(button, "Close");
            cross.BindTheme(Avalonia.Controls.Shapes.Shape.StrokeProperty, nameof(AppTheme.TextFaint));
            button.PointerEntered += (s, e) =>
            {
                button.BindTheme(Border.BackgroundProperty, nameof(AppTheme.ButtonHover));
                cross.BindTheme(Avalonia.Controls.Shapes.Shape.StrokeProperty, nameof(AppTheme.Text));
            };
            button.PointerExited += (s, e) =>
            {
                button.BindTheme(Border.BackgroundProperty, key: null, whenNone: Brushes.Transparent);
                cross.BindTheme(Avalonia.Controls.Shapes.Shape.StrokeProperty, nameof(AppTheme.TextFaint));
            };
            button.PointerReleased += (s, e) =>
            {
                if (e.InitialPressMouseButton == Avalonia.Input.MouseButton.Left)
                {
                    CloseSidePanel();
                }
            };
            return button;
        }

        public bool IsSidePanelShown
        {
            get
            {
                return sidePanel.IsVisible;
            }
        }

        /// <summary>
        /// Puts the side panel away (its cross, a click on empty chrome). A tool's own panel can be
        /// the only way on with that tool (Function, the angle of a rotation, the dialogs of
        /// Define figure), so closing it puts the tool down too: back to Drag, a construction in
        /// progress abandoned, as with Escape; picking the tool again brings the panel back.
        /// Anything else (a figure's properties, the length panel) is just hidden, and comes back
        /// with a click on the figure. A tool's panel is told by its type: some tools make a new
        /// one each time they are asked.
        /// </summary>
        public void CloseSidePanel()
        {
            var drawing = CurrentDrawing;
            var shown = PropertyGrid.Selection;
            var behavior = drawing?.Behavior;
            if (shown != null
                && behavior != null
                && behavior != Behavior.Default
                && behavior.PropertyBag?.GetType() == shown.GetType())
            {
                drawing.SetDefaultBehavior();
            }

            ShowProperties(null);
        }

        public void ToggleLabelNewPoints()
        {
            Settings.Instance.AutoLabelPoints = !Settings.Instance.AutoLabelPoints;
        }

        public void ToggleGrid()
        {
            // the grid is the drawing's (a file says whether it shows), so this is an undo
            // step - except in the middle of a construction, whose step is its figure's alone
            var grid = CurrentDrawing.CoordinateGrid;
            if (CurrentDrawing.IsRecordingTransaction)
            {
                grid.Visible = !grid.Visible;
            }
            else
            {
                CurrentDrawing.ActionManager.SetProperty(grid, nameof(grid.Visible), !grid.Visible);
            }

            // the G key comes through here too, and the ribbon button must follow
            CommandToolButton.UpdateToggles();
        }

        public void ToggleFigureExplorer()
        {
            ShowFigureExplorer(!FigureExplorer.IsVisible);
        }

        /// <summary>The Figure List in the first column, as wide as it was last, or the column closed</summary>
        void ShowFigureExplorer(bool visible)
        {
            var column = ColumnDefinitions[0];
            // (not laid out yet at startup)
            if (!visible && FigureExplorer.IsVisible && column.ActualWidth > 0)
            {
                figureExplorerWidth = column.ActualWidth;
            }

            Settings.Instance.ShowFigureExplorer = visible;
            FigureExplorer.IsVisible = visible;
            figureExplorerSplitter.IsVisible = visible;
            column.MinWidth = visible ? 120 : 0;
            column.Width = new GridLength(visible ? figureExplorerWidth : 0);
            if (visible)
            {
                FigureExplorer.Refresh();
            }

            CommandToolButton.UpdateToggles();
        }

        public void ToggleOrtho()
        {
            Settings.Instance.EnableOrtho = !Settings.Instance.EnableOrtho;
            Settings.Instance.EnablePolar = false;
        }

        public void TogglePolar()
        {
            Settings.Instance.EnablePolar = !Settings.Instance.EnablePolar;
            Settings.Instance.EnableOrtho = false;
        }

        public void ToggleSnapToGrid()
        {
            Settings.Instance.EnableSnapToGrid = !Settings.Instance.EnableSnapToGrid;
        }

        public void ToggleSnapToPoint()
        {
            Settings.Instance.EnableSnapToPoint = !Settings.Instance.EnableSnapToPoint;
        }

        public void ToggleSnapToCenter()
        {
            Settings.Instance.EnableSnapToCenter = !Settings.Instance.EnableSnapToCenter;
        }

        public void TogglePointByCoordinates()
        {
            Settings.Instance.EnablePointByCoordinates = !Settings.Instance.EnablePointByCoordinates;

            // the panel of the current tool appears or goes right away
            if (CurrentDrawing != null && CurrentDrawing.Behavior != null)
            {
                ShowProperties(CurrentDrawing.Behavior.PropertyBag);
            }
        }

        protected void CreateCanvas()
        {
            DrawingControl = new DrawingControl();
            DrawingControl.HorizontalAlignment = HorizontalAlignment.Stretch;
            DrawingControl.VerticalAlignment = VerticalAlignment.Stretch;
            DrawingControl.ReadyForInteraction += RaiseReadyForInteraction;
            DrawingControl.DrawingAttach += DrawingControl_DrawingAttach;
            DrawingControl.DrawingDetach += DrawingControl_DrawingDetach;
        }

        protected void RaiseReadyForInteraction(object sender, EventArgs e)
        {
            if (ReadyForInteraction != null)
            {
                ReadyForInteraction(sender, e);
            }
        }

        public virtual void RaiseCommandExecuted(Command command)
        {
            // Do nothing when a command is executed but allow this to be overridden.
        }

        protected virtual void DrawingControl_DrawingAttach(Drawing drawing)
        {
            drawing.Status += mCurrentDrawing_Status;
            drawing.ChoiceStatus += mCurrentDrawing_ChoiceStatus;
            drawing.SelectionChanged += mCurrentDrawing_SelectionChanged;
            drawing.BehaviorChanged += mCurrentDrawing_BehaviorChanged;
            drawing.DisplayProperties += mCurrentDrawing_DisplayProperties;
            drawing.UnhandledException += UnhandledException;
            drawing.FigureCoordinatesChanged += mCurrentDrawing_FigureCoordinatesChanged;
            FigureExplorer.Drawing = drawing;
        }

        protected virtual void DrawingControl_DrawingDetach(Drawing drawing)
        {
            drawing.Status -= mCurrentDrawing_Status;
            drawing.ChoiceStatus -= mCurrentDrawing_ChoiceStatus;
            mCurrentDrawing_ChoiceStatus(null);
            drawing.SelectionChanged -= mCurrentDrawing_SelectionChanged;
            drawing.BehaviorChanged -= mCurrentDrawing_BehaviorChanged;
            drawing.DisplayProperties -= mCurrentDrawing_DisplayProperties;
            drawing.UnhandledException -= UnhandledException;
            drawing.FigureCoordinatesChanged -= mCurrentDrawing_FigureCoordinatesChanged;
            FigureExplorer.Drawing = null;
            ShowProperties(null);
        }

        public BehaviorToolButton AddToolButton(Behavior behavior)
        {
            return Ribbon.AddToolButton(behavior);
        }

        /// <param name="first">Before the tools of its tab instead of after them</param>
        public CommandToolButton AddToolbarButton(Command command, bool first = false)
        {
            return Ribbon.AddToolButton(command, first);
        }

        public void RemoveToolButton(Behavior behavior)
        {
            Ribbon.RemoveToolButton(behavior);
        }

        protected virtual void Behavior_NewBehaviorCreated(Behavior behavior)
        {
            AddToolButton(behavior);
        }

        protected virtual void Behavior_BehaviorDeleted(Behavior behavior)
        {
            RemoveToolButton(behavior);
        }

        public void CreateRibbon()
        {
            Ribbon = new Ribbon(this);
        }

        public void AddBehaviors(Assembly assembly)
        {
            var behaviors = Behavior.LoadBehaviors(assembly);
            foreach (var behavior in behaviors)
            {
                AddToolButton(behavior);
            }
        }

        public void Clear()
        {
            this.DrawingControl.Clear();
        }

        protected virtual void mCurrentDrawing_DisplayProperties(object sender, Drawing.DisplayPropertiesEventArgs e)
        {
            ShowProperties(e.Object, e.FocusProperty);
        }

        protected virtual void mCurrentDrawing_BehaviorChanged(Behavior newBehavior)
        {
            Ribbon.SelectBehavior(newBehavior);
            var help = newBehavior.HintText;
            if (!help.IsEmpty())
            {
                ShowHint(help);
            }
            ShowProperties(newBehavior.PropertyBag);
        }

        protected virtual void mCurrentDrawing_SelectionChanged(object sender, Drawing.SelectionChangedEventArgs e)
        {
            ShowSelectionProperties();
        }

        // a segment's length in the grid follows a drag of its end: at once, then at most
        // every 300 ms while the drag goes on, the last position always included (undoing a
        // point drag replays every mouse step of it)
        void mCurrentDrawing_FigureCoordinatesChanged(object sender, Drawing.FigureCoordinatesChangedEventArgs e)
        {
            Throttle.Schedule(
                PropertyGrid,
                grid => Avalonia.Threading.Dispatcher.UIThread.Post(grid.RefreshNumbers),
                TimeSpan.FromMilliseconds(300),
                ThrottleOptions.RunOnceImmediatelyIfFree);
        }

        private void mCurrentDrawing_Status(string status)
        {
            ShowHint(status);
        }

        // the hint, and what a click at the cursor would take while it could take more than
        // one thing (ClickChoice), which is shown over the hint until it goes
        string hint;
        string choiceHint;

        void mCurrentDrawing_ChoiceStatus(string text)
        {
            choiceHint = text;
            UpdateStatusBar();
        }

        public virtual void ShowHint(string text)
        {
            hint = text;
            UpdateStatusBar();
        }

        void UpdateStatusBar()
        {
            if (Settings.Instance.HideHints)
            {
                return;
            }

            var text = choiceHint ?? hint;
            if (text.IsEmpty())
            {
                StatusBar.Visibility = Visibility.Collapsed;
            }
            else
            {
                StatusBar.Text = text;
                StatusBar.Visibility = Visibility.Visible;
            }
        }

        protected virtual void ShowSelectionProperties()
        {
            var selection = CurrentDrawing.GetSelectedFigures().ToArray();
            if (selection.Length == 1)
            {
                ShowProperties(selection[0]);
            }
            else if (selection.Length > 1)
            {
                ShowProperties(new FigureSelection(CurrentDrawing, selection, PropertyGrid.ValueDiscoveryStrategy));
            }
            else
            {
                ShowProperties(null);
            }
        }

        /// <param name="focusProperty">The property whose editor takes the keyboard; null for none</param>
        public virtual void ShowProperties(object selection, string focusProperty = null)
        {
            try
            {
                PropertyGrid.Show(selection, CurrentDrawing.ActionManager, focusProperty);
            }
            catch (Exception ex)
            {
                CurrentDrawing.RaiseError(this, ex);
            }
        }

        public void ShowProperties(IEnumerable<object> selection)
        {
            try
            {
                PropertyGrid.Show(selection, CurrentDrawing.ActionManager);
            }
            catch (Exception ex)
            {
                CurrentDrawing.RaiseError(this, ex);
            }
        }
    }
}
