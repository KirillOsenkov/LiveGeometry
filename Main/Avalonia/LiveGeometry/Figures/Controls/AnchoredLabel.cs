using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// A label hung on one figure, its first dependency: a point's name (<see cref="PointLabel"/>),
/// a line's or a circle's name (<see cref="FigureLabel"/>) and a text attached to a figure
/// (<see cref="AttachedLabel"/>). It sits at its offset from an anchor on the figure, and a
/// new one at the figure's default place (<see cref="FigureAnchor"/>); it keeps clear of other
/// names (<see cref="NameLayout"/>), shows only while its figure does, and beside a line, which
/// runs past the window, stays on screen.
/// </summary>
public abstract class AnchoredLabel : LabelWithOffset
{
    /// <summary>The figure the label hangs on</summary>
    public IFigure Figure => Dependencies.Count > 0 ? Dependencies[0] : null;

    /// <summary>Where the offset is measured from: on the figure, and for a line where the eye finds it</summary>
    public override Point Anchor => FigureAnchor.Of(Figure, Drawing?.CoordinateSystem);

    protected override bool KeepsClear => true;

    public override IFigure OwnPoint => Figure as IPoint;

    /// <summary>
    /// A point's lines and circles - what it is built on or on, and what is built on it; a
    /// segment, a line or a circle itself: a segment's name doesn't go across its segment
    /// </summary>
    public override IEnumerable<IFigure> FiguresThroughAnchor()
    {
        var figure = Figure;
        if (figure == null)
        {
            return Enumerable.Empty<IFigure>();
        }

        return figure is IPoint ? figure.Dependents.Concat(figure.Dependencies) : new[] { figure };
    }

    protected override ZOrder DefaultLayer()
    {
        return ZOrder.PointLabels;
    }

    protected override Point? DefaultOffset()
    {
        return FigureAnchor.DefaultOffset(Figure, MeasureSize(), Drawing.CoordinateSystem);
    }

    /// <summary>Whether the label goes while its figure is hidden or doesn't exist (a point's name doesn't: <see cref="PointLabel.HidesWithFigure"/>)</summary>
    protected virtual bool HidesWithFigure => true;

    public override void UpdateVisual()
    {
        var figure = Figure;
        if (figure == null || Drawing == null)
        {
            return;
        }

        base.UpdateVisual();
        if (!HidesWithFigure)
        {
            return;
        }

        bool shown = Visible && figure.Visible && figure.Exists;
        if (figure is LineBase && figure is not Segment)
        {
            // a line that runs outside the window has nowhere to write it
            shown = shown && FigureAnchor.IsOnScreen(ToPhysical(Anchor), Drawing.CoordinateSystem);
            if (shown)
            {
                KeepOnScreen();
            }
        }

        // a hidden figure keeps its label to itself
        Shape.IsVisible = shown;
    }

    /// <summary>The label pushed back in, whole, where a line's anchor by the edge of the window and the offset put it past the edge (<see cref="FigureAnchor.KeptOnScreen"/>)</summary>
    void KeepOnScreen()
    {
        var topLeft = ToPhysical(Coordinates);
        var kept = FigureAnchor.KeptOnScreen(topLeft, MeasureSize(), Drawing.CoordinateSystem);
        if (kept != topLeft && kept.Exists())
        {
            Coordinates = ToLogical(kept);
            Shape.MoveTo(kept);
        }
    }
}
