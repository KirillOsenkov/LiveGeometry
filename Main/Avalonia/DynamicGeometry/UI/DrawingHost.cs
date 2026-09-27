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

        public Command CommandToggleGrid { get; set; }
        public Command CommandToggleOrtho { get; set; }
        public Command CommandToggleSnapToGrid { get; set; }
        public Command CommandToggleSnapToPoint { get; set; }
        public Command CommandToggleLabelNewPoints { get; set; }
        public Command CommandTogglePolar { get; set; }
        public Command CommandToggleSnapToCenter { get; set; }
        public Command CommandTogglePointByCoordinates { get; set; }
        public Command CommandDrawingBackground { get; set; }
        public Command CommandToggleFigureExplorer { get; set; }

        GridSplitter figureExplorerSplitter;

        /// <summary>The width the Figure List had when it was put away, to come back with</summary>
        double figureExplorerWidth = 240;

        public DrawingHost()
        {
            Behavior.NewBehaviorCreated += Behavior_NewBehaviorCreated;
            Behavior.BehaviorDeleted += Behavior_BehaviorDeleted;
            SetupLayout();
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
            this.Children.Add(propertyGridScrollViewer);
            this.Children.Add(StatusBar);

            Grid.SetColumnSpan(Ribbon, 3);
            Grid.SetRow(FigureExplorer, 1);
            Grid.SetRow(figureExplorerSplitter, 1);
            Grid.SetColumn(figureExplorerSplitter, 1);
            foreach (var control in new Control[] { DrawingControl, propertyGridScrollViewer, StatusBar })
            {
                Grid.SetRow(control, 1);
                Grid.SetColumn(control, 2);
            }

            ShowFigureExplorer(Settings.Instance.ShowFigureExplorer);

            CommandToggleGrid = new Command(ToggleGrid, CartesianGrid.GetIcon(), "Grid", BehaviorCategories.Coordinates)
            {
                IsChecked = () => CurrentDrawing != null && CurrentDrawing.CoordinateGrid.Visible,
                Shortcut = "G"
            };
            CommandToggleOrtho = new Command(ToggleOrtho, ToggleIcons.Ortho(), "Ortho", BehaviorCategories.Selection)
            {
                IsChecked = () => Settings.Instance.EnableOrtho
            };
            CommandToggleSnapToGrid = new Command(ToggleSnapToGrid, ToggleIcons.SnapToGrid(), "Snap to grid", BehaviorCategories.Selection)
            {
                IsChecked = () => Settings.Instance.EnableSnapToGrid
            };
            CommandToggleSnapToPoint = new Command(ToggleSnapToPoint, ToggleIcons.SnapToPoint(), "Snap to point", BehaviorCategories.Selection)
            {
                IsChecked = () => Settings.Instance.EnableSnapToPoint
            };
            CommandToggleLabelNewPoints = new Command(ToggleLabelNewPoints, ToggleIcons.LabelNewPoints(), "Label new points", BehaviorCategories.Points)
            {
                IsChecked = () => Settings.Instance.AutoLabelPoints
            };
            CommandTogglePolar = new Command(TogglePolar, ToggleIcons.Polar(), "Polar", BehaviorCategories.Selection)
            {
                IsChecked = () => Settings.Instance.EnablePolar
            };
            CommandToggleSnapToCenter = new Command(ToggleSnapToCenter, ToggleIcons.SnapToCenter(), "Snap to center", BehaviorCategories.Selection)
            {
                IsChecked = () => Settings.Instance.EnableSnapToCenter
            };
            CommandTogglePointByCoordinates = new Command(TogglePointByCoordinates, ToggleIcons.PointByCoordinates(), "Point by coordinates", BehaviorCategories.Coordinates)
            {
                IsChecked = () => Settings.Instance.EnablePointByCoordinates
            };
            CommandDrawingBackground = new Command(ToggleDrawingProperties, ToggleIcons.Background(), "Background", BehaviorCategories.Coordinates);
            CommandToggleFigureExplorer = new Command(ToggleFigureExplorer, ToggleIcons.FigureList(), "Figure List", BehaviorCategories.Selection)
            {
                IsChecked = () => FigureExplorer.IsVisible
            };
        }

        /// <summary>The drawing's own properties (its paper) in the side panel; again to put them away</summary>
        public void ToggleDrawingProperties()
        {
            if (CurrentDrawing == null)
            {
                return;
            }

            ShowProperties(PropertyGrid.Selection == CurrentDrawing ? null : CurrentDrawing);
        }

        protected void CreateFigureExplorer()
        {
            FigureExplorer = new FigureExplorer();
            figureExplorerSplitter = new GridSplitter()
            {
                Width = 4,
                ResizeDirection = GridResizeDirection.Columns,
                Background = RibbonTheme.HeaderRowBackground
            };
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
            propertyGridScrollViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                Margin = new Thickness(8),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                MinWidth = 200.0,
                Visibility = Visibility.Collapsed
            };

            PropertyGrid = new PropertyGrid();

            // the same surface as the tools strip of the ribbon
            propertyGridScrollViewer.Content = new Border()
            {
                Background = RibbonTheme.Background,
                BorderBrush = RibbonTheme.TabLine,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(14, 10, 14, 12),
                Child = PropertyGrid
            };
            PropertyGrid.VisibilityChanged += PropertyGrid_VisibilityChanged;

            propertyGridScrollViewer.ZIndex = (int)ZOrder.StatusBar;

            PropertyGrid.ValueDiscoveryStrategy = new ExcludeByDefaultValueDiscoveryStrategy();
        }

        private void PropertyGrid_VisibilityChanged(object sender, EventArgs e)
        {
            propertyGridScrollViewer.Visibility = PropertyGrid.Visibility;
        }

        public void ToggleLabelNewPoints()
        {
            Settings.Instance.AutoLabelPoints = !Settings.Instance.AutoLabelPoints;
        }

        public void ToggleGrid()
        {
            CurrentDrawing.CoordinateGrid.Visible = !CurrentDrawing.CoordinateGrid.Visible;

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
            ShowProperties(e.Object);
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

        public virtual void ShowHint(string text)
        {
            if (!Settings.Instance.HideHints)
            {
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

        public virtual void ShowProperties(object selection)
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
