using GuiLabs.Undo;

namespace LiveGeometry;

public interface IValueEditor
{
    IValueProvider Value { get; set; }
    ActionManager ActionManager { get; set; }
}
