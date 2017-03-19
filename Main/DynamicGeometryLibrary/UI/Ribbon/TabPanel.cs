using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace DynamicGeometry
{
    public class TabPanel : TabItem
    {

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

        protected override void OnSelected(System.Windows.RoutedEventArgs e)
        {
            base.OnSelected(e);
            if (Settings.UpdateSelectedBehaviorOnTabChange)
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
