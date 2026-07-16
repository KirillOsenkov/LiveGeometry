using System.Reflection;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public interface IValueEditor
    {
        IValueProvider Value { get; set; }
        ActionManager ActionManager { get; set; }
    }
}
