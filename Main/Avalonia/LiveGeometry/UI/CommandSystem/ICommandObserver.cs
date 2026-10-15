namespace LiveGeometry;

public interface ICommandObserver
{
    void CommandRemoved();
    void EnabledChanged(bool newEnabledState);
    void IconChanged(FrameworkElement icon);
}
