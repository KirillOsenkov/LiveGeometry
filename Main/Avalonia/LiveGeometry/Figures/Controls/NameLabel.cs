namespace LiveGeometry;

/// <summary>
/// The name written next to a figure: a point's, with or without its coordinates
/// (<see cref="PointLabel"/>), or a line's, a ray's, a segment's or a circle's
/// (<see cref="FigureLabel"/>). Part of its figure (<see cref="FigureBase.NameLabel"/>), not a
/// figure of its own: shown and hidden by the figure's Show name, which adds and removes it,
/// and left out of the Figure List, Delete, a region's selection, a show/hide box and a
/// formula, which take the figure instead.
/// </summary>
public abstract class NameLabel : AnchoredLabel
{
    /// <summary>"of A", "of line g": "Label of A", "Name of line g"</summary>
    public override string Construction => Figure != null ? "of " + ConstructionText.Of(Figure) : null;

    public override void OnAddingToDrawing(Drawing drawing)
    {
        base.OnAddingToDrawing(drawing);

        // undo puts the label back without going through the figure's setter
        if (Figure is FigureBase figure && figure.NameLabel == null)
        {
            figure.NameLabel = this;
        }
    }

    public override void OnRemovingFromDrawing(Drawing drawing)
    {
        base.OnRemovingFromDrawing(drawing);
        if (Figure is FigureBase figure && figure.NameLabel == this)
        {
            figure.NameLabel = null;
        }
    }

    /// <summary>
    /// Not in the grid: the name shows and hides with its figure's Show name. (Hidden by
    /// itself, the label stayed hidden whatever Show name said, and a name label isn't in the
    /// Figure List to be found again.)
    /// </summary>
    [PropertyGridVisible(false)]
    public override bool Visible
    {
        get => base.Visible; set => base.Visible = value;
    }

    /// <summary>Not in the grid: what a name says is worked out from its figure</summary>
    [PropertyGridVisible(false)]
    public override string Text
    {
        get => base.Text; set => base.Text = value;
    }
}
