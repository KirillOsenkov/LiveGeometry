using GuiLabs.Undo;

namespace DynamicGeometry
{
    public interface ISupportRemoveDependency
    {
        bool CanRemoveDependency(IFigure dependency);
        IAction GetRemoveDependencyAction(IFigure dependency);
    }
}
