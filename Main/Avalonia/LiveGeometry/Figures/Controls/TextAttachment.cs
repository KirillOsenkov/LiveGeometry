using GuiLabs.Undo;

namespace LiveGeometry;

/// <summary>
/// A text attached to a figure (<see cref="AttachedLabel"/>) and a free one (<see cref="Label"/>),
/// one turned into the other where it is drawn (2026-10-17): Detach on an attached text,
/// and a drag with Alt onto a figure or onto the paper (<see cref="Dragger"/>), as a point
/// dragged with Alt snaps onto a figure or lets go of it. The new text takes the old one's
/// text, style, visibility, lock and selection; one undo step (the drag's, when dropped).
/// </summary>
public static class TextAttachment
{
    /// <summary>What a text can be attached to: a point, a segment, a line, a ray, a vector, a circle or a polygon</summary>
    public static bool CanAttachTo(IFigure figure)
    {
        return figure is IPoint or ILine or ICircle or PolygonBase;
    }

    /// <summary>
    /// Whether the text can go onto a figure: one attached already, or a free one on the
    /// paper (not pinned to the screen) whose text is plain - an attached text shows [...]
    /// parts as typed, so a free one with them stays free
    /// </summary>
    public static bool CanAttach(IFigure text)
    {
        return text is AttachedLabel
            || text is Label { Pin: LabelPin.None } free && LabelBase.FindExpressionParts(free.Text ?? "").Count == 0;
    }

    /// <summary>Why a free text can't be attached, for the status; null when it can or isn't one</summary>
    public static string WhyNot(IFigure text)
    {
        return text is Label { Pin: LabelPin.None } && !CanAttach(text)
            ? "A text with [...] parts can't be attached to a figure: it would show them as typed."
            : null;
    }

    /// <summary>The text attached to the figure, where it is drawn now</summary>
    public static AttachedLabel Attach(LabelBase text, IFigure figure)
    {
        var drawing = text.Drawing;
        var attached = Factory.CreateAttachedLabel(drawing, figure);
        attached.Text = text.Text;
        CopyLook(text, attached);
        attached.PlaceAt(text.Coordinates);
        Replace(text, attached);
        return attached;
    }

    /// <summary>The text let go of its figure: a free one where it is drawn now</summary>
    public static Label Detach(AttachedLabel attached)
    {
        var free = Factory.CreateLabel(attached.Drawing);
        free.Text = attached.Text;
        CopyLook(attached, free);
        free.MoveTo(attached.Coordinates);
        Replace(attached, free);
        return free;
    }

    static void CopyLook(LabelBase from, LabelBase to)
    {
        to.Style = from.Style;
        to.Visible = from.Visible;
        to.Locked = from.Locked;
    }

    static void Replace(LabelBase old, LabelBase replacement)
    {
        var drawing = old.Drawing;
        bool selected = old.Selected;
        using (Transaction.Create(drawing.ActionManager, false))
        {
            Actions.Remove(old);
            Actions.Add(drawing, replacement);
        }

        // (outside the step, as a paste selects its copies: the halo goes with the text)
        if (selected)
        {
            drawing.Figures.ClearSelection();
            replacement.Selected = true;
            drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        }
    }
}
