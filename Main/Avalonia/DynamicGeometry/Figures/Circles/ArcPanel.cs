using System.Collections.Generic;

namespace DynamicGeometry;

/// <summary>
/// The side panel of an arc, a sector or a circular segment right after it is made or
/// converted: the two verbs that make the other two kinds of it, and OK. An arc, a sector
/// and a segment are the same three clicks (center, start, end), so the shapes have no tool
/// of their own: the arc is drawn and this panel finishes it (<see cref="CircleArcCreator.ShowCreatedFigure"/>),
/// and a verb clicked here ends in this panel for the new figure, so that one more click
/// goes on to the third kind or back (the same verb in the figure's own grid ends in the
/// new figure's grid: <see cref="EllipseArc.Convert"/>). The verb also trades the figure's
/// style for the filled or unfilled one of its hue (<see cref="StyleManager.ConvertStyle"/>).
/// Same for the elliptical kinds.
/// </summary>
public class ArcPanel : ICustomPropertyProvider, ICustomMethodProvider
{
    readonly EllipseArcBase figure;

    public ArcPanel(EllipseArcBase figure)
    {
        this.figure = figure;
    }

    /// <summary>Puts the panel for the figure up, and says so in the status bar</summary>
    public static void Show(EllipseArcBase figure)
    {
        var drawing = figure.Drawing;
        var panel = new ArcPanel(figure);
        drawing.RaiseDisplayProperties(panel);
        drawing.RaiseStatusNotification(figure.Title + ": make it " + panel.OtherKinds + " in the panel, or go on.");
    }

    bool IsSector => figure is CircleSector || figure is EllipseSector;

    bool IsSegment => figure is CircleSegment || figure is EllipseSegment;

    bool IsArc => !IsSector && !IsSegment;

    /// <summary>"a sector or a segment": the two kinds the verbs make</summary>
    string OtherKinds => IsArc ? "a sector or a segment"
        : IsSector ? "an arc or a segment"
        : "an arc or a sector";

    /// <summary>No rows, only the verbs</summary>
    public IEnumerable<IValueProvider> GetProperties()
    {
        yield break;
    }

    /// <summary>
    /// The verbs to the other two kinds, in the order arc, sector, segment; not to an arc
    /// while something measures the shape's area or perimeter, as the figure's own grid
    /// (an arc has neither)
    /// </summary>
    public IEnumerable<IOperationDescription> GetMethods()
    {
        if (!IsArc && !figure.IsUsedForArea())
        {
            yield return MethodDescription.Get<ArcPanel>(nameof(ConvertToArc));
        }

        if (!IsSector)
        {
            yield return MethodDescription.Get<ArcPanel>(nameof(ConvertToSector));
        }

        if (!IsSegment)
        {
            yield return MethodDescription.Get<ArcPanel>(nameof(ConvertToSegment));
        }

        yield return MethodDescription.Get<ArcPanel>(nameof(OK));
    }

    [PropertyGridName("Convert to arc")]
    [PropertyGridIcon(PropertyGridIcon.Arc)]
    public void ConvertToArc()
    {
        Convert(figure is CircleArcBase
            ? Factory.CreateArc(figure.Drawing, figure.Dependencies)
            : Factory.CreateEllipseArc(figure.Drawing, figure.Dependencies));
    }

    [PropertyGridName("Convert to sector")]
    [PropertyGridIcon(PropertyGridIcon.Sector)]
    public void ConvertToSector()
    {
        Convert(figure is CircleArcBase
            ? Factory.CreateCircleSector(figure.Drawing, figure.Dependencies)
            : Factory.CreateEllipseSector(figure.Drawing, figure.Dependencies));
    }

    [PropertyGridName("Convert to segment")]
    [PropertyGridIcon(PropertyGridIcon.CircleSegment)]
    public void ConvertToSegment()
    {
        Convert(figure is CircleArcBase
            ? Factory.CreateCircleSegment(figure.Drawing, figure.Dependencies)
            : Factory.CreateEllipseSegment(figure.Drawing, figure.Dependencies));
    }

    /// <summary>The new figure takes the old one's place (one undo step) and gets this panel in turn</summary>
    void Convert(IArc replacement)
    {
        EllipseArc.Replace(figure, replacement);
        Show((EllipseArcBase)replacement);
    }

    /// <summary>Closes the panel; the figure stays as it is</summary>
    [PropertyGridIcon(PropertyGridIcon.Check)]
    public void OK()
    {
        figure.Drawing.RaiseDisplayProperties(null);
        figure.Drawing.ShowBehaviorHint();
    }

    // the title of the panel
    public override string ToString()
    {
        return figure.Title;
    }
}
