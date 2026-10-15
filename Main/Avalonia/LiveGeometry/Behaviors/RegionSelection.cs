using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using AvaloniaShapes = Avalonia.Controls.Shapes;

namespace LiveGeometry;

/// <summary>
/// A selection by a region under way (<see cref="SelectionRegion"/>): the Rectangle and Lasso
/// selection tools, and a drag with the right button under the Drag tool. The region is drawn
/// as a band of the selection's stripes, and what it takes is selected while it grows, so the
/// halos show what the release selects. It adds to the selection, or with Alt (read at every
/// move) takes out of it. <see cref="Cancel"/> (Escape) puts back what was selected. A selection
/// is no undo step, so nothing is recorded.
/// </summary>
public class RegionSelection
{
    readonly Drawing drawing;
    readonly SelectionRegion region;

    // what was selected before: figures, and parts selected by themselves
    readonly HashSet<IFigure> selectedBefore;
    readonly AvaloniaShapes.Path boundary;
    readonly bool showsKeyHint;
    HashSet<IFigure> taken = new HashSet<IFigure>();
    bool removes;

    /// <param name="showsKeyHint">Whether the status says what Alt does (not to a finger, which has no keys to hold)</param>
    public RegionSelection(Drawing drawing, SelectionRegion region, bool showsKeyHint)
    {
        this.drawing = drawing;
        this.region = region;
        this.showsKeyHint = showsKeyHint;
        selectedBefore = new HashSet<IFigure>(drawing.GetSelectedFigures());
        boundary = new AvaloniaShapes.Path()
        {
            Stroke = SelectionHalo.Brush,
            StrokeThickness = SelectionHalo.BandWidth,
            StrokeJoin = PenLineJoin.Round,
            StrokeLineCap = PenLineCap.Round,
            IsHitTestVisible = false
        };

        // with the halos, which it looks like: under the points, and as faint
        SelectionHalo.Layer.Get(drawing.Canvas).Children.Add(boundary);
        Update(removing: false);
        ShowStatus();
    }

    /// <summary>The pointer is at the point now; <paramref name="removing"/>: Alt is held</summary>
    public void Extend(Point point, bool removing)
    {
        if (region.Extend(point) || removing != removes)
        {
            Update(removing);
        }
    }

    void Update(bool removing)
    {
        boundary.Data = region.CreateGeometry();
        var now = new HashSet<IFigure>(region.FindFigures(drawing));

        // the figures whose state changes: those taken or let go of since, or all those
        // taken now and before when Alt went down or up
        var changed = removing != removes
            ? taken.Union(now).ToList()
            : taken.Except(now).Union(now.Except(taken)).ToList();
        if (removing != removes)
        {
            removes = removing;
            ShowStatus();
        }

        taken = now;
        foreach (var figure in changed)
        {
            Apply(figure);
        }
    }

    void Apply(IFigure figure)
    {
        if (!taken.Contains(figure))
        {
            Restore(figure);
        }
        else if (removes)
        {
            figure.Selected = false;
        }
        else
        {
            figure.Selected = true;
        }
    }

    /// <summary>The figure selected as it was before, a vertex or side selected by itself included</summary>
    void Restore(IFigure figure)
    {
        figure.Selected = selectedBefore.Contains(figure);
        foreach (var part in FigureParts.WithParts(figure).Skip(1))
        {
            if (selectedBefore.Contains(part))
            {
                part.Selected = true;
            }
        }
    }

    void ShowStatus()
    {
        if (showsKeyHint)
        {
            drawing.RaiseStatusNotification(removes
                ? "Release to unselect the figures. Let go of " + KeyNames.Alt + " to select them instead."
                : "Release to select the figures. Hold " + KeyNames.Alt + " to unselect them instead.");
        }
    }

    /// <summary>The release: what the region took stays selected (or unselected)</summary>
    public void Finish()
    {
        RemoveBoundary();
        drawing.ShowBehaviorHint();
        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
    }

    /// <summary>The selection is put back as it was before</summary>
    public void Cancel()
    {
        RemoveBoundary();
        var released = taken.ToList();
        taken.Clear();
        foreach (var figure in released)
        {
            Restore(figure);
        }

        drawing.ShowBehaviorHint();
    }

    void RemoveBoundary()
    {
        (boundary.Parent as Panel)?.Children.Remove(boundary);
    }
}
