using System.ComponentModel;

namespace DynamicGeometry;

/// <summary>
/// The Point tool with the X/Y panel always on: a point typed by its coordinates, without
/// first switching "Point by coordinates" on in the Coordinates tab. That toggle stays for
/// typing the points of other figures (a segment's ends, a circle's center); this tool is
/// the short way to a point alone. A click on the canvas still makes a point as the Point
/// tool does.
/// </summary>
[Category(BehaviorCategories.Points)]
[Order(4)]
public class PointByCoordinatesCreator : FreePointCreator
{
    protected override bool ShowsCoordinatesPanel
    {
        get { return true; }
    }

    public override FrameworkElement CreateIcon()
    {
        return IconBuilder.BuildIcon()
            .Point(0.14, 0.5)
            .Text(nameof(AppTheme.Ink), 0.38, 0.06, text: "X =", fontSize: 11)
            .Text(nameof(AppTheme.Ink), 0.38, 0.5, text: "Y =", fontSize: 11)
            .Canvas;
    }

    public override string Name
    {
        get { return "Coordinates"; }
    }

    public override string HintText
    {
        get { return "Type the coordinates of the new point in the panel, or click to create a point."; }
    }
}
