using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

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
        public Command CommandShowFigureExplorer { get; set; }

        public DrawingHost()
        {
            Behavior.NewBehaviorCreated += Behavior_NewBehaviorCreated;
            Behavior.BehaviorDeleted += Behavior_BehaviorDeleted;
            SetupLayout();
        }

        protected virtual void SetupLayout()
        {
            this.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
            this.RowDefinitions.Add(new RowDefinition());
            this.ColumnDefinitions.Add(new ColumnDefinition());
            this.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });

            CreateRibbon();
            CreateCanvas();
            CreatePropertyGrid();
            CreateStatusBar();
            CreateFigureExplorer();

            this.Children.Add(Ribbon);
            this.Children.Add(DrawingControl);
            this.Children.Add(propertyGridScrollViewer);
            this.Children.Add(StatusBar);
            this.Children.Add(FigureExplorer);

            FigureExplorer.Visible = Settings.Instance.ShowFigureExplorer;

            Grid.SetColumnSpan(Ribbon, 2);
            Grid.SetColumn(FigureExplorer, 1);
            Grid.SetRow(FigureExplorer, 1);
            Grid.SetRow(DrawingControl, 1);
            Grid.SetRow(propertyGridScrollViewer, 1);
            Grid.SetRow(StatusBar, 1);

            CommandToggleGrid = new Command(ToggleGrid, CartesianGrid.GetIcon(), "Grid", BehaviorCategories.Coordinates);
            CommandToggleOrtho = new Command(ToggleOrtho, new CheckBox(), "Ortho", BehaviorCategories.Selection);
            CommandToggleSnapToGrid = new Command(ToggleSnapToGrid, new CheckBox(), "Snap to grid", BehaviorCategories.Selection);
            CommandToggleSnapToPoint = new Command(ToggleSnapToPoint, new CheckBox(), "Snap to point", BehaviorCategories.Selection);
            CommandToggleLabelNewPoints = new Command(ToggleLabelNewPoints, new CheckBox(), "Label New Points", BehaviorCategories.Points);
            CommandTogglePolar = new Command(TogglePolar, new CheckBox(), "Polar", BehaviorCategories.Selection);
            CommandToggleSnapToCenter = new Command(ToggleSnapToCenter, new CheckBox(), "Snap to Center", BehaviorCategories.Selection);
            CommandShowFigureExplorer = new Command(ToggleFigureExplorer, new CheckBox() { IsChecked = FigureExplorer.Visible }, "Figure List", BehaviorCategories.Drawing);
        }

        protected void CreateFigureExplorer()
        {
            FigureExplorer = new FigureExplorer()
            {
                MinWidth = 200,
                MaxWidth = 400
            };
            FigureExplorer.SelectionChanged += FigureExplorer_SelectionChanged;
        }

        bool guard = false; // to prevent reentrancy
        void FigureExplorer_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (guard)
            {
                return;
            }

            guard = true;

            foreach (var deselected in e.RemovedItems)
            {
                IFigure deselectedFigure = deselected as IFigure;
                if (deselectedFigure != null)
                {
                    deselectedFigure.Selected = false;
                }
            }

            foreach (var selected in e.AddedItems)
            {
                IFigure selectedFigure = selected as IFigure;
                if (selectedFigure != null)
                {
                    selectedFigure.Selected = true;
                }
            }

            CurrentDrawing.RaiseSelectionChanged(CurrentDrawing.GetSelectedFigures());

            guard = false;
        }

        void drawing_SelectionChanged(object sender, Drawing.SelectionChangedEventArgs e)
        {
            SyncFigureExplorerSelection();
        }

        private void SyncFigureExplorerSelection()
        {
            if (guard)
            {
                return;
            }
            guard = true;

            // Temporary Solution.  This causes figure's name change to show in FigureExplorer. - D.H.
            // Same temporary solution is used in ToggleFigureExplorer.
            if (FigureExplorer.Visible)
            {
                FigureExplorer.ItemsSource = null;
                FigureExplorer.ItemsSource = CurrentDrawing.Figures;
            }
            // End Temporary Solution

            FigureExplorer.SelectedItem = null;
            foreach (var selectedFigure in CurrentDrawing.GetSelectedFigures())
            {
                FigureExplorer.SelectedItems.Add(selectedFigure);
            }
            guard = false;
        }

        protected virtual void CreateStatusBar()
        {
            StatusBar = new StatusBar();
            StatusBar.HorizontalAlignment = HorizontalAlignment.Left;
            StatusBar.VerticalAlignment = VerticalAlignment.Bottom;
            Canvas.SetZIndex(StatusBar, (int)ZOrder.StatusBar);
        }

        protected void CreatePropertyGrid()
        {
            propertyGridScrollViewer = new ScrollViewer()
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Margin = new Thickness(8),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                MinWidth = 200.0,
                Background = new SolidColorBrush(Color.FromArgb(255, 255, 255, 233)),
                Visibility = Visibility.Collapsed
            };

            PropertyGrid = new PropertyGrid();
            PropertyGrid.Margin = new Thickness(4);
            propertyGridScrollViewer.Content = PropertyGrid;
            PropertyGrid.VisibilityChanged += PropertyGrid_VisibilityChanged;

            Canvas.SetZIndex(propertyGridScrollViewer, (int)ZOrder.StatusBar);

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
        }

        public void ToggleFigureExplorer()
        {
            if (FigureExplorer.Visible)
            {
                FigureExplorer.Visible = false;
            }
            else
            {
                // Temporary Solution.  This causes figure's name change to should show in FigureExplorer. - D.H.
                // Same temporary solution is used in SyncFigureExplorerSelection().
                FigureExplorer.ItemsSource = null;
                FigureExplorer.ItemsSource = CurrentDrawing.Figures;
                FigureExplorer.Visible = true;
            }
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
            drawing.SelectionChanged += drawing_SelectionChanged;
            FigureExplorer.ItemsSource = drawing.Figures;
        }

        protected virtual void DrawingControl_DrawingDetach(Drawing drawing)
        {
            drawing.Status -= mCurrentDrawing_Status;
            drawing.SelectionChanged -= mCurrentDrawing_SelectionChanged;
            drawing.BehaviorChanged -= mCurrentDrawing_BehaviorChanged;
            drawing.DisplayProperties -= mCurrentDrawing_DisplayProperties;
            drawing.UnhandledException -= UnhandledException;
            drawing.SelectionChanged -= drawing_SelectionChanged;
            FigureExplorer.ItemsSource = null;
            ShowProperties(null);
        }

        public BehaviorToolButton AddToolButton(Behavior behavior)
        {
            return Ribbon.AddToolButton(behavior);
        }

        public CommandToolButton AddToolbarButton(Command command)
        {
            return Ribbon.AddToolButton(command);
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
                ShowProperties(selection);
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
