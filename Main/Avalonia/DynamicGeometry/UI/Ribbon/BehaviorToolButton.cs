using System.ComponentModel;

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

            var shortcut = BehaviorShortcuts.GetShortcut(behavior);
            var tip = behavior.Name + (shortcut != null ? "  (" + shortcut + ")" : "");
            if (!string.IsNullOrEmpty(behavior.HintText))
            {
                tip += "\n" + behavior.HintText;
            }

            Avalonia.Controls.ToolTip.SetTip(this, new Avalonia.Controls.TextBlock()
            {
                Text = tip,
                MaxWidth = 320,
                TextWrapping = Avalonia.Media.TextWrapping.Wrap
            });
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

            // The tool that is on already (its tab was opened, which picks it): the press took
            // the keyboard out of its panel's box, and what is typed next would pick tools by
            // their letters. Shown again, the box takes the keyboard back.
            bool isCurrent = DrawingHost.CurrentDrawing.Behavior == ParentBehavior;
            DrawingHost.CurrentDrawing.Behavior = ParentBehavior;
            ParentPanel.SelectedToolButton = this;
            if (isCurrent && ParentBehavior.PropertyBag != null)
            {
                DrawingHost.ShowProperties(ParentBehavior.PropertyBag);
            }

            // and so is its hint, which what the tool made since had replaced (a new segment's
            // "set its length"); not halfway through a construction, where the status says
            // what to click next
            if (isCurrent && ParentBehavior.IsInInitialState && !ParentBehavior.HintText.IsEmpty())
            {
                DrawingHost.ShowHint(ParentBehavior.HintText);
            }
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