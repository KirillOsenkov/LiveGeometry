using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Controls.Shapes;
using Avalonia.Media;

namespace LiveGeometry;

/// <summary>
/// The three tools that make a <see cref="Region"/>: Difference, Intersection and Union.
/// Each click takes a figure with an inside (a click inside it, or on its edge; the
/// smallest of nested ones first, Tab for the next); Enter, a right click, a click on the
/// paper or a click on a shape taken already makes the region of what was clicked, once
/// there are two. The hover shows the shape a click would take, and from the first click
/// on, what the region would be, tinted. For Difference the first click is the shape and
/// the rest is cut out of it; the other two take their shapes in any order. Once there are
/// two, the side panel says what the region would read as and offers Done
/// (<see cref="RegionPanel"/>), so that the way to finish is in sight.
/// </summary>
public abstract class RegionCreator : FigureCreator
{
    protected abstract RegionOperation Operation { get; }

    /// <summary>
    /// The tool's panel while a region can be made of what was clicked: the shapes so far, as
    /// the region will read, and Done. A new one at every step, so that the row says the
    /// latest shape. None before the second shape, none once the region is made
    /// ("construction complete" re-reads it before the found figures are cleared, but after
    /// the transaction is gone).
    /// </summary>
    public override object PropertyBag => Transaction != null && CanFinish ? new RegionPanel(this) : null;

    [PropertyGridNoUndo]
    public class RegionPanel
    {
        readonly RegionCreator tool;

        public RegionPanel(RegionCreator tool)
        {
            this.tool = tool;
        }

        /// <summary>The shapes clicked so far, as the region will read: "square ABCD without circle c"</summary>
        [PropertyGridVisible]
        [PropertyGridName("Region")]
        public string Shapes => Region.Describe(tool.Operation, tool.FoundDependencies);

        /// <summary>Makes the region of the shapes clicked so far</summary>
        [PropertyGridVisible]
        [PropertyGridIcon(PropertyGridIcon.Check)]
        public void Done()
        {
            tool.TryFinish();
        }

        /// <summary>The panel's title: the tool's name</summary>
        public override string ToString()
        {
            return tool.Name;
        }
    }

    protected override DependencyList InitExpectedDependencies()
    {
        // not used: every step wants a figure with an inside (GetExpectedDependencyType)
        return DependencyList.Create<Polygon, Polygon>();
    }

    protected override Type GetExpectedDependencyType()
    {
        return typeof(IRegion);
    }

    protected override bool ExpectingAPoint()
    {
        return false;
    }

    protected override bool CanCreateTempResults()
    {
        return false;
    }

    protected override void AddFoundDependency(IFigure figure)
    {
        if (figure is IRegion)
        {
            FoundDependencies.Add(figure);
        }
    }

    /// <summary>
    /// What a click here takes: the figures with an inside containing the place, the
    /// smallest first (the inner one of nested shapes), then those hit on their edge; not
    /// one taken already, unless a click on it finishes the region
    /// </summary>
    protected override IReadOnlyList<IFigure> FindExpectedDependencies(Point coordinates)
    {
        if (Drawing == null)
        {
            return Array.Empty<IFigure>();
        }

        var inside = PointInRegion.FindRegions(Drawing, coordinates).Cast<IFigure>();
        var onEdge = Drawing.Figures
            .HitTestAll(coordinates, f => FigureParts.Whole(f) is IRegion region && f.Visible && f.IsHitTestVisible && region.HasInside)
            .Select(FigureParts.Whole);
        return inside
            .Concat(onEdge)
            .Distinct()
            .Where(f => !FoundDependencies.Contains(f) || FinishesOn(f))
            .ToList();
    }

    /// <summary>Two shapes or more: enough for a region</summary>
    bool CanFinish => FoundDependencies.Count >= 2;

    /// <summary>A click on a shape taken already finishes the region, once there are two</summary>
    protected override bool FinishesOn(IFigure figure)
    {
        return figure != null && CanFinish && FoundDependencies.Contains(figure);
    }

    protected override string FinishStatus => "Click here again, press Done or Enter to make the region.";

    /// <summary>A click on the paper, or on a shape taken already, finishes; any other goes on</summary>
    protected override void Click(Point coordinates)
    {
        var under = LookForExpectedDependencyUnderCursor(ClickedUnconstrainedCoordinates);
        if (CanFinish && (under == null || FinishesOn(under)))
        {
            TryFinish();
            return;
        }

        base.Click(coordinates);
    }

    bool TryFinish()
    {
        if (!CanFinish)
        {
            return false;
        }

        RemoveIntermediateFigureIfNecessary();
        AddFiguresAndRestart();
        return true;
    }

    public override void MouseRightClick(object sender, MouseButtonEventArgs e)
    {
        if (TryFinish())
        {
            return;
        }

        base.MouseRightClick(sender, e);
    }

    public override void KeyDown(object sender, Avalonia.Input.KeyEventArgs e)
    {
        if (e.Key == Avalonia.Input.Key.Enter && TryFinish())
        {
            e.Handled = true;
            return;
        }

        base.KeyDown(sender, e);
    }

    protected override IEnumerable<IFigure> CreateFigures()
    {
        yield return Factory.CreateRegion(Drawing, Operation, FoundDependencies.ToList());
    }

    /// <summary>The region's page: its area and perimeter are right there, and the operation can be changed</summary>
    protected override void ShowCreatedFigure(IList<IFigure> figures)
    {
        var region = figures.OfType<Region>().FirstOrDefault();
        if (region != null)
        {
            Drawing.RaiseDisplayProperties(region);
        }
    }

    /// <summary>What the region would be of the shapes clicked so far and the one under the cursor, tinted</summary>
    protected override RegionBoundary GetRegionPreview(MouseEventArgs e)
    {
        if (FoundDependencies.Count == 0)
        {
            return null;
        }

        var operands = FoundDependencies.OfType<IRegion>().ToList();
        if (GetFigureToPick(e) is IRegion hovered && !operands.Contains(hovered))
        {
            operands.Add(hovered);
        }

        return RegionBoundary.Trace(Operation, operands);
    }

    public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
    {
        if (FoundDependencies.Count == 0)
        {
            return HintText;
        }

        if (FoundDependencies.Count == 1)
        {
            return NextShapeHint(FoundDependencies[0]);
        }

        return "Click more shapes, or press Done (or Enter) to make the region.";
    }

    /// <summary>What the second click takes, with the first shape named</summary>
    protected abstract string NextShapeHint(IFigure first);

    protected override bool CanPlacePointOn(IFigure figure)
    {
        return false;
    }

    /// <summary>
    /// Two overlapping circles outlined in ink, the part the operation keeps filled as a
    /// shape is: the Venn diagram of the operation
    /// </summary>
    protected static FrameworkElement CreateVennIcon(RegionOperation operation)
    {
        var builder = IconBuilder.BuildIcon();
        double size = builder.Canvas.Width;
        const double radius = 0.28;
        const double leftX = 0.36;
        const double rightX = 0.64;
        var left = new EllipseGeometry(new Rect(size * (leftX - radius), size * (0.5 - radius), size * 2 * radius, size * 2 * radius));
        var right = new EllipseGeometry(new Rect(size * (rightX - radius), size * (0.5 - radius), size * 2 * radius, size * 2 * radius));
        var mode = operation == RegionOperation.Difference ? GeometryCombineMode.Exclude
            : operation == RegionOperation.Intersection ? GeometryCombineMode.Intersect
            : GeometryCombineMode.Union;
        var fill = new Path() { Data = new CombinedGeometry(mode, left, right) };
        fill.BindTheme(Shape.FillProperty, nameof(AppTheme.ShapeIconFill));
        builder.Canvas.Children.Add(fill);
        builder
            .Circle(leftX, 0.5, radius)
            .Circle(rightX, 0.5, radius);
        return builder.Canvas;
    }
}

[Category(BehaviorCategories.Shapes)]
[Order(7)]
public class RegionDifferenceCreator : RegionCreator
{
    protected override RegionOperation Operation => RegionOperation.Difference;

    public override string Name => "Difference";

    public override string HintText => "Click a shape, then the shapes to cut out of it. Enter when done.";

    protected override string NextShapeHint(IFigure first)
    {
        return "Now click the shapes to cut out of " + ConstructionText.Of(first) + ".";
    }

    public override FrameworkElement CreateIcon()
    {
        return CreateVennIcon(RegionOperation.Difference);
    }
}

[Category(BehaviorCategories.Shapes)]
[Order(8)]
public class RegionIntersectionCreator : RegionCreator
{
    protected override RegionOperation Operation => RegionOperation.Intersection;

    public override string Name => "Intersection";

    public override string HintText => "Click the shapes that overlap: the region is what they share. Enter when done.";

    protected override string NextShapeHint(IFigure first)
    {
        return "Now click the shapes that overlap " + ConstructionText.Of(first) + ".";
    }

    public override FrameworkElement CreateIcon()
    {
        return CreateVennIcon(RegionOperation.Intersection);
    }
}

[Category(BehaviorCategories.Shapes)]
[Order(9)]
public class RegionUnionCreator : RegionCreator
{
    protected override RegionOperation Operation => RegionOperation.Union;

    public override string Name => "Union";

    public override string HintText => "Click the shapes to join into one region. Enter when done.";

    protected override string NextShapeHint(IFigure first)
    {
        return "Now click the shapes to join with " + ConstructionText.Of(first) + ".";
    }

    public override FrameworkElement CreateIcon()
    {
        return CreateVennIcon(RegionOperation.Union);
    }
}
