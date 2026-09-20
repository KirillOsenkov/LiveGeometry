using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace DynamicGeometry
{
    public class Ribbon : TabControl
    {
        // Avalonia matches theme templates by concrete type; keep using the TabControl template.
        protected override System.Type StyleKeyOverride => typeof(TabControl);

        public DrawingHost DrawingHost { get; set; }

        public Ribbon(DrawingHost drawingHost)
        {
            DrawingHost = drawingHost;
            Background = RibbonTheme.Background;
            BorderBrush = RibbonTheme.BottomBorder;
            BorderThickness = new Thickness(0, 0, 0, 1);
            Padding = new Thickness(4, 0, 4, 3);
        }

        public BehaviorToolButton AddToolButton(Behavior behavior)
        {
            BehaviorToolButton button = behavior.CreateToolButton();
            string category = BehaviorOrderer.GetCategory(behavior);
            AddToolButton(button, category);
            return button;
        }

        public void SelectBehavior(Behavior behavior)
        {
            var panel = FindPanel(behavior);
            SelectedItem = panel;
            var button = panel.FindButton(behavior);
            if (button != null)
            {
                button.Click();
            }
        }

        public BehaviorToolButton FindButton(Behavior behavior)
        {
            var panel = FindPanel(behavior);
            var button = panel.FindButton(behavior);
            return button;
        }

        public TabPanel FindPanel(Behavior behavior)
        {
            foreach (var panel in Panels)
            {
                BehaviorToolButton button = panel.FindButton(behavior);
                if (button != null)
                {
                    return panel;
                }
            }
            return null;
        }

        private IEnumerable<TabPanel> Panels
        {
            get
            {
                return Items.OfType<TabPanel>();
            }
        }

        public TabPanel GetTabPanelByCategory(ToolButton button, string category)
        {
            var result = GetPanel(category);

            if (result == null)
            {
                result = CreateTabPanel(button, category);
            }

            return result;
        }

        private TabPanel CreateTabPanel(ToolButton button, string category)
        {
            var result = new TabPanel()
            {
                Category = category,
                Panel = new WrapPanel(),
                HeaderContent = new ButtonGrid(Settings.ShowIconInTabPanelHeader ? button.CloneIcon() : null, category, true),
                MinHeight = 34,
                Padding = new Thickness(9, 0, 9, 0)
            };
            Items.Add(result);
            return result;
        }

        public CommandToolButton AddToolButton(Command command)
        {
            var button = new CommandToolButton(command);
            AddToolButton(button, command.Category);
            return button;
        }

        public void AddToolButton(ToolButton button, string category)
        {
            button.DrawingHost = DrawingHost;
            var panel = GetTabPanelByCategory(button, category);

            // a thin divider between the tools of a tab and the option toggles that follow
            var last = panel.Panel.Children.LastOrDefault();
            if (button is CommandToolButton && last is BehaviorToolButton)
            {
                panel.Panel.Children.Add(new Avalonia.Controls.Shapes.Rectangle()
                {
                    Width = 1,
                    Height = 44, // a WrapPanel doesn't stretch its children
                    Fill = RibbonTheme.Separator,
                    Margin = new Thickness(6, 0, 6, 0),
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
                });
            }

            panel.Panel.Children.Add(button);
            button.ParentPanel = panel;
        }

        public void RemoveToolButton(Behavior behavior)
        {
            var panel = FindPanel(behavior);
            var button = FindButton(behavior);
            button.CommandRemoved();
            panel.ResetSelectedToolButton();
        }

        public TabPanel GetPanel(string category)
        {
            var result = Items
                .OfType<TabPanel>()
                .Where(p => p.Category == category)
                .FirstOrDefault();
            return result;
        }
    }
}
