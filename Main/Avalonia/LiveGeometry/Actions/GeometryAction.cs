using GuiLabs.Undo;

namespace LiveGeometry;

public abstract class GeometryAction : AbstractAction
{
    public GeometryAction(Drawing drawing)
    {
        Drawing = drawing;
    }

    public Drawing Drawing { get; set; }
}
