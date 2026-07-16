using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;

namespace DynamicGeometry
{
    public interface ICommandObserver
    {
        void CommandRemoved();
        void EnabledChanged(bool newEnabledState);
        void IconChanged(FrameworkElement icon);
    }
}
