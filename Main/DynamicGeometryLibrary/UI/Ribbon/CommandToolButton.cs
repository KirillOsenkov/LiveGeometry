using System.Windows.Controls;

namespace DynamicGeometry
{
    public class CommandToolButton : ToolButton
    {
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
            buttonGrid.MouseLeftButtonDown += Content_MouseLeftButtonDown;
        }

        private void Content_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Click();
            ToggleCheckBox();
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