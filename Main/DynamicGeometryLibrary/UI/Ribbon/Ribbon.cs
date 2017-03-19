using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DynamicGeometry
{
    public class Ribbon : TabControl
    {
        public DrawingHost DrawingHost { get; set; }

        public Ribbon(DrawingHost drawingHost)
        {
            DrawingHost = drawingHost;
            Background = new LinearGradientBrush()
            {
                StartPoint = new Point(0, 0),
                EndPoint = new Point(0, 1),
                GradientStops = new GradientStopCollection()
                {
                    new GradientStop()
                    {
                        Offset = 0.9,
                        Color = Colors.White
                    },
                    new GradientStop()
                    {
                        Offset = 1,
                        Color = Color.FromArgb(255, 230, 230, 230)
                    }
                }
            };
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
                HeaderContent = (Settings.ShowIconInTabPanelHeader) ? new ButtonGrid(button.CloneIcon(), category) : new ButtonGrid(null, category)
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
