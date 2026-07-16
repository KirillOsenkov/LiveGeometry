using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;

namespace DynamicGeometry
{
    public class BehaviorToolButton : ToolButton
    {
        public static double IconTextGap = 0;
        public BehaviorToolButton(Behavior behavior)
        {
            ParentBehavior = behavior;
            behavior.PropertyChanged += behavior_PropertyChanged;

            buttonGrid = new ButtonGrid(behavior.Icon, behavior.Name);
            buttonGrid.IconTextGap = IconTextGap;
            Content = buttonGrid;
            buttonGrid.PointerPressed += buttonGrid_MouseLeftButtonDown;
        }

        public override FrameworkElement CloneIcon()
        {
            return ParentBehavior.CreateIcon();
        }

        public bool IsChecked
        {
            get
            {
                return buttonGrid.IsChecked;
            }
            set
            {
                buttonGrid.IsChecked = value;
            }
        }

        void buttonGrid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Click();
        }

        public override void Click()
        {
            if (DrawingHost.CurrentDrawing == null)
            {
                return;
            }
            DrawingHost.CurrentDrawing.Behavior = ParentBehavior;
            ParentPanel.SelectedToolButton = this;
        }

        void behavior_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Name")
            {
                buttonGrid.Text = ParentBehavior.Name;
            }
        }

        public Behavior ParentBehavior { get; set; }
    }
}