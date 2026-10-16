using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia.Input;

namespace LiveGeometry;

/// <summary>
/// One click inside a polygon, a circle, a sector or a path puts a <see cref="PointInRegion"/>
/// there: a point that is dragged like a free one but never leaves the figure. Offered
/// whether or not the figure is filled (the hover's halo on the figure says which one);
/// where regions nest, the smallest first and Tab for the next (<see cref="ClickChoice"/>).
/// A click on a point, or inside nothing, does nothing.
/// </summary>
[Category(BehaviorCategories.Points)]
[Order(4)]
public class PointInRegionCreator : Behavior
{
    PointPlacement hoverPlacement;

    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        var placement = Choice.Pick(FindPlacements(e));
        if (placement != null)
        {
            Actions.Add(Drawing, placement.Create(Drawing));
        }
    }

    /// <summary>A point inside each region containing the click, the smallest region first; none over a point</summary>
    List<PointPlacement> FindPlacements(MouseEventArgs e)
    {
        var coordinates = Coordinates(e);
        if (Drawing.Figures.HitTest<IPoint>(Coordinates(e, false, false, false)) != null)
        {
            return new List<PointPlacement>();
        }

        return PointInRegion.FindRegions(Drawing, coordinates)
            .Select(region => PointPlacement.InRegion(region, coordinates))
            .ToList();
    }

    protected override IReadOnlyList<object> FindClickOptions(MouseEventArgs e)
    {
        return FindPlacements(e);
    }

    public override void MouseMove(object sender, MouseEventArgs e)
    {
        hoverPlacement = Choice.Pick(FindPlacements(e));
    }

    protected override PointPlacement GetClickPreview(MouseEventArgs e)
    {
        return hoverPlacement;
    }

    protected override Cursor GetCursor(Avalonia.Point coordinates)
    {
        return GetCursor(hoverPlacement);
    }

    public override FrameworkElement CreateIcon()
    {
        return IconBuilder.BuildIcon()
            .Polygon(
                nameof(AppTheme.ShapeIconFill),
                nameof(AppTheme.ShapeOutline),
                new Avalonia.Point(0.1, 0.9),
                new Avalonia.Point(0.5, 0.1),
                new Avalonia.Point(0.9, 0.9))
            .Point(0.5, 0.62, nameof(AppTheme.PointOnFigureFill))
            .Canvas;
    }

    public override string Name => "Point inside";

    public override string HintText => "Click inside a polygon, circle, sector or path to create a point that stays inside it.";
}
