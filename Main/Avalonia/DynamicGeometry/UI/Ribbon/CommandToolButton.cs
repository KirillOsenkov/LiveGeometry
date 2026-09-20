using System.Collections.Generic;
using Avalonia.Controls;

namespace DynamicGeometry
{
    public class CommandToolButton : ToolButton
    {
        /// <summary>
        /// Toggles can switch each other off (Ortho and Polar), so after any of them
        /// is clicked all of them re-read their state.
        /// </summary>
        static readonly List<CommandToolButton> toggles = new List<CommandToolButton>();

        public CommandToolButton(Command command)
        {
            Command = command;
            command.AddObserver(this);

            buttonGrid = new ButtonGrid(command.Icon, command.Name);
            if (command.Icon is CheckBox)
            {
                command.Icon.IsHitTestVisible = false;
            }
            Content = buttonGrid;
            buttonGrid.PointerPressed += Content_MouseLeftButtonDown;

            if (command.IsChecked != null)
            {
                toggles.Add(this);
                AttachedToVisualTree += (s, e) => UpdateCheckedState();
            }
        }

        private void Content_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Click();
            ToggleCheckBox();
            foreach (var toggle in toggles)
            {
                toggle.UpdateCheckedState();
            }
        }

        void UpdateCheckedState()
        {
            buttonGrid.IsChecked = Command.IsChecked();
        }

        private void ToggleCheckBox()
        {
            CheckBox check = Command.Icon as CheckBox;
            if (check != null)
            {
                check.IsChecked = !check.IsChecked;
            }
        }

        public override void Click()
        {
            Command.Execute();
        }

        public Command Command { get; set; }

        public override void EnabledChanged(bool newEnabledState)
        {
            base.EnabledChanged(newEnabledState);
            buttonGrid.Opacity = (newEnabledState) ? 1.0 : 0.5;
        }
    }
}
