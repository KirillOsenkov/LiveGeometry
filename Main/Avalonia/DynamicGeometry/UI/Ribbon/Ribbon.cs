using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Controls.Templates;

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
            Template = CreateTemplate();
        }

        /// <summary>
        /// Our own template instead of the theme's: a row of group headers with a line along
        /// its bottom, and the tools of the selected group below. The line is behind the
        /// headers, so that the selected one (see <see cref="TabOutline"/>) can paint over it
        /// and open into the tools.
        /// </summary>
        readonly Border headerStartHost = new Border()
        {
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
        };

        readonly Border headerEndHost = new Border()
        {
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            IsHitTestVisible = false
        };

        /// <summary>
        /// Goes before the group headers, in the same row (the application puts its document
        /// buttons there: one strip less than a toolbar of their own). Always fully visible;
        /// in a narrow window it is the headers that run out of room.
        /// </summary>
        public Control HeaderStart
        {
            get
            {
                return headerStartHost.Child;
            }
            set
            {
                headerStartHost.Child = value;
            }
        }

        /// <summary>
        /// Something unimportant for the far right of the header row (the build stamp).
        /// It is behind the headers, which cover it when the window is narrow.
        /// </summary>
        public Control HeaderEnd
        {
            get
            {
                return headerEndHost.Child;
            }
            set
            {
                headerEndHost.Child = value;
            }
        }

        IControlTemplate CreateTemplate()
        {
            return new FuncControlTemplate<TabControl>((tabControl, scope) =>
            {
                var headers = new ItemsPresenter()
                {
                    Name = "PART_ItemsPresenter",
                    // the headers overlap each other, so the first one sticks out to the left
                    Margin = new Thickness(6 + ButtonGrid.HeaderOverlap, 4, 6 + ButtonGrid.HeaderOverlap, 0),
                    [!ItemsPresenter.ItemsPanelProperty] = tabControl[!ItemsPanelProperty]
                }.RegisterInNameScope(scope);

                var headerRow = new Panel() { Background = RibbonTheme.HeaderRowBackground };
                headerRow.Children.Add(new Border()
                {
                    BorderBrush = RibbonTheme.TabLine,
                    BorderThickness = new Thickness(0, 0, 0, 1)
                });
                headerRow.Children.Add(headerEndHost);

                var startAndHeaders = new DockPanel();
                DockPanel.SetDock(headerStartHost, Dock.Left);
                startAndHeaders.Children.Add(headerStartHost);
                startAndHeaders.Children.Add(headers);
                headerRow.Children.Add(startAndHeaders);
                DockPanel.SetDock(headerRow, Dock.Top);

                var tools = new Border()
                {
                    Background = RibbonTheme.Background,
                    BorderBrush = RibbonTheme.BottomBorder,
                    BorderThickness = new Thickness(0, 0, 0, 1),
                    Padding = new Thickness(6, 3, 6, 3),
                    Child = new ContentPresenter()
                    {
                        Name = "PART_SelectedContentHost",
                        [!ContentPresenter.ContentProperty] = tabControl[!SelectedContentProperty],
                        [!ContentPresenter.ContentTemplateProperty] = tabControl[!SelectedContentTemplateProperty]
                    }.RegisterInNameScope(scope)
                };

                var root = new DockPanel();
                root.Children.Add(headerRow);
                root.Children.Add(tools);
                return root;
            });
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
                HeaderContent = new ButtonGrid(Settings.ShowIconInTabPanelHeader ? button.CloneIcon() : null, category, isTabHeader: true)
            };
            Items.Add(result);
            return result;
        }

        public CommandToolButton AddToolButton(Command command, bool first = false)
        {
            var button = new CommandToolButton(command);
            if (first)
            {
                AddFirstToolButton(button, command.Category);
            }
            else
            {
                AddToolButton(button, command.Category);
            }

            return button;
        }

        /// <summary>
        /// A command that leads its tab (the grid switch under Coordinates): before the tools,
        /// after the leading commands already there, with one thin divider between the run of
        /// them and the tools.
        /// </summary>
        void AddFirstToolButton(ToolButton button, string category)
        {
            button.DrawingHost = DrawingHost;
            var panel = GetTabPanelByCategory(button, category);
            if (panel.LeadingCount == 0 && panel.Panel.Children.Count > 0)
            {
                panel.Panel.Children.Insert(0, CreateDivider());
            }

            panel.Panel.Children.Insert(panel.LeadingCount, button);
            panel.LeadingCount++;
            button.ParentPanel = panel;
        }

        static Control CreateDivider()
        {
            return new Avalonia.Controls.Shapes.Rectangle()
            {
                Width = 1,
                Height = 44, // a WrapPanel doesn't stretch its children
                Fill = RibbonTheme.Separator,
                Margin = new Thickness(6, 0, 6, 0),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
            };
        }

        public void AddToolButton(ToolButton button, string category)
        {
            button.DrawingHost = DrawingHost;
            var panel = GetTabPanelByCategory(button, category);

            // a thin divider between the tools of a tab and the option toggles that follow
            var last = panel.Panel.Children.LastOrDefault();
            if (button is CommandToolButton && last is BehaviorToolButton)
            {
                panel.Panel.Children.Add(CreateDivider());
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
