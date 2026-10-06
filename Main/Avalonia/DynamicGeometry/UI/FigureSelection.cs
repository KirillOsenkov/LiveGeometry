using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry;

/// <summary>
/// Several figures in the property grid: the properties they have in common, and Delete for
/// all of them at once, as the Delete key does. Two Bezier paths or more can be made one
/// with holes.
/// </summary>
public class FigureSelection : CompositePropertyProvider, IConditionalProperties
{
    readonly Drawing drawing;
    readonly IFigure[] figures;

    public FigureSelection(Drawing drawing, IEnumerable<IFigure> figures, IValueDiscoveryStrategy valueDiscoveryStrategy)
        : base(valueDiscoveryStrategy, figures)
    {
        this.drawing = drawing;
        this.figures = figures.ToArray();
    }

    /// <summary>The largest of the Bezier paths selected leaves the others out of its inside (<see cref="BezierPath.CutHoles"/>)</summary>
    [PropertyGridVisible]
    [PropertyGridName("Cut out holes")]
    [PropertyGridIcon(PropertyGridIcon.Polygon)]
    public void CutHoles()
    {
        BezierPath.CutHoles(drawing, figures);
    }

    [PropertyGridVisible]
    [PropertyGridName("Delete")]
    [PropertyGridDestructive]
    public void Delete()
    {
        drawing.Delete(figures);
    }

    public bool CanEdit(string propertyName)
    {
        return propertyName != nameof(CutHoles) || BezierPath.CanCutHoles(figures);
    }

    public string Caption(string propertyName, string defaultCaption)
    {
        return defaultCaption;
    }

    public override string ToString()
    {
        return figures.Length + " figures selected";
    }
}
