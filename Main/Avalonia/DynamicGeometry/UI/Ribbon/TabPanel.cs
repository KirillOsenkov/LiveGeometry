using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;

namespace DynamicGeometry
{
    public class TabPanel : TabItem
    {
        // Avalonia matches theme templates by concrete type; keep using the TabItem template.
        protected override System.Type StyleKeyOverride => typeof(TabItem);

        public string Category { get; set; }

        BehaviorToolButton selectedToolButton;
        public BehaviorToolButton SelectedToolButton
        {
            get
            {
                return selectedToolButton;
            }
            set
            {
                if (selectedToolButton == value)
                {
                    return;
                }
                if (selectedToolButton != null)
                {
                    selectedToolButton.IsChecked = false;
                }
                selectedToolButton = value;
                if (selectedToolButton != null)
                {
                    selectedToolButton.IsChecked = true;
                    if (Settings.ShowIconInTabPanelHeader)
                    {
                        HeaderContent.Icon = selectedToolButton.CloneIcon();
                    }
                }
            }
        }

        protected override void OnPropertyChanged(Avalonia.AvaloniaPropertyChangedEventArgs change)
        {
            base.OnPropertyChanged(change);
            if (change.Property == IsSelectedProperty
                && (bool)change.NewValue
                && Settings.UpdateSelectedBehaviorOnTabChange)
            {
                UpdateSelectedToolButton();
            }
        }

        public void ResetSelectedToolButton()
        {
            SelectedToolButton = null;
            UpdateSelectedToolButton();
        }

        public void UpdateSelectedToolButton()
        {
            if (selectedToolButton == null)
            {
                var first = BehaviorToolButtons.FirstOrDefault();
                if (first != null)
                {
                    first.Click();
                }
            }
            else
            {
                selectedToolButton.Click();
            }
        }

        WrapPanel panel;
        public WrapPanel Panel
        {
            get
            {
                return panel;
            }
            set
            {
                panel = value;
                Content = value;
            }
        }

        IEnumerable<BehaviorToolButton> BehaviorToolButtons
        {
            get
            {
                return Panel.Children.OfType<BehaviorToolButton>();
            }
        }

        ButtonGrid headerContent;
        public ButtonGrid HeaderContent
        {
            get
            {
                return headerContent;
            }
            set
            {
                headerContent = value;
                Header = value;
            }
        }

        public BehaviorToolButton FindButton(Behavior behavior)
        {
            return BehaviorToolButtons.FirstOrDefault(t => t.ParentBehavior == behavior);
        }
    }
}
