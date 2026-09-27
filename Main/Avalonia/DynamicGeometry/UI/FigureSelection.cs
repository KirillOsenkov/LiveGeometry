using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry;

/// <summary>
/// Several figures in the property grid: the properties they have in common, and Delete for
/// all of them at once, as the Delete key does.
/// </summary>
public class FigureSelection : CompositePropertyProvider
{
    readonly Drawing drawing;
    readonly IFigure[] figures;

    public FigureSelection(Drawing drawing, IEnumerable<IFigure> figures, IValueDiscoveryStrategy valueDiscoveryStrategy)
        : base(valueDiscoveryStrategy, figures)
    {
        this.drawing = drawing;
        this.figures = figures.ToArray();
    }

    [PropertyGridVisible]
    [PropertyGridName("Delete")]
    [PropertyGridDestructive]
    public void Delete()
    {
        drawing.Delete(figures);
    }

    public override string ToString()
    {
        return figures.Length + " figures selected";
    }
}
