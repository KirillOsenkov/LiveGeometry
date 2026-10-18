using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Input;
using Avalonia.Media;
using GuiLabs.Undo;

namespace LiveGeometry;

[Category(BehaviorCategories.Selection)]
[Order(1)]
public partial class Dragger : Behavior
{
    /// <summary>
    /// What one press holds while it lasts: the figure under it, what the drag moves,
    /// where it started. One per pointer (the mouse is one, each finger on a figure is
    /// one, see "Touch" in <see cref="Behavior"/>): two fingers drag the two ends of a
    /// segment at once, each its own.
    /// </summary>
    class Drag
    {
        public IPointer Pointer;
        public List<IMovable> Moving = new List<IMovable>();
        public IFigure Found;

        // A label the press found that can't be dragged: the drag moves the view, a
        // click selects it (a show/hide box: ticks it)
        public ControlBase PressedFixedLabel;
        public List<IFigure> ToRecalculate;
        public Point OffsetFromFigureLeftTopCorner;
        public Point OldCoordinates;
        public Point CoordinatesOnMouseDown;
        public bool StartedMoving;

        // the press was on a figure selected with others, and the drag moves them all
        public bool DraggingSelection;

        // Where the point dragged with Alt goes if dropped now; null while it is free
        public PointPlacement Snap;

        // A text dragged with Alt: the figure it goes onto if dropped now, or whether an
        // attached one lets go of its figure (over the paper)
        public IFigure AttachTarget;
        public bool DetachesText;
    }

    // the presses held right now, in the order they came
    readonly List<Drag> drags = new List<Drag>();

    // A drag is one undo step, with the releases and the snap of an Alt-drag in it -
    // one for all the fingers down at once, from the first drag that moves something
    // to the last release (transactions nest, they can't be interleaved)
    Transaction dragTransaction;

    /// <summary>In cursor tolerances: how far a snapped point may be pulled before it lets go</summary>
    public static double StickyReach = 2;

    /// <summary>In pixels: how far the cursor goes from where it was pressed before the press is a drag and not a click</summary>
    public static double DragThreshold = 3;

    public Dragger()
    {
        // the status says what is under the cursor also when it is one thing
        Choice.DescribesSingle = true;
    }

    Drag FindDrag(IPointer pointer)
    {
        return drags.Find(drag => drag.Pointer == pointer);
    }

    /// <summary>
    /// What a press at the coordinates takes: the figure a drag moves or a click
    /// selects, or null for the paper, which a drag moves - an axis (the grid's, for the
    /// tools to build on), and the labels that can't be dragged
    /// (<see cref="Drawing.FixedLabels"/>); a drag that starts on a caption takes the
    /// captions along (<paramref name="isCaption"/>). A click still selects such a
    /// label (<paramref name="fixedLabel"/>), and one that is selected drags as any
    /// label does - but a caption, where on a phone a tap and then a thumb to read on
    /// would pull the explanation away from its heading. Nor does a tap select the
    /// caption: the side panel would cover half of a phone for a reader touching the
    /// text.
    /// </summary>
    IFigure FindPressedFigure(Point coordinates, bool isTouch, out ControlBase fixedLabel, out bool isCaption)
    {
        fixedLabel = null;
        isCaption = false;
        var found = Drawing.Figures.HitTest(coordinates);

        // the figure chosen with Tab among those under the cursor (FindClickOptions)
        if (Choice.Index > 0 && Choice.Current is IFigure chosen)
        {
            found = chosen;
        }

        if (found is AxisLine)
        {
            return null;
        }

        if (found is ControlBase label && Drawing.FixedLabels.Contains(label))
        {
            isCaption = found is Label { Pin: not LabelPin.None };
            if (isCaption || !label.Selected)
            {
                if (!isCaption || !isTouch)
                {
                    fixedLabel = label;
                }

                return null;
            }
        }

        return found;
    }

    /// <summary>
    /// A finger stays with what its press would drag or select, whatever the other
    /// fingers do; one whose press would move the view (the paper, an axis, a caption,
    /// a graph with nothing to move) pans and zooms with the other paper fingers
    /// </summary>
    protected override bool TakesFingerOnFigure(MouseButtonEventArgs e)
    {
        var drag = PrepareDrag(e);
        return !drag.Moving.Contains(Drawing.CoordinateSystem);
    }

    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        // a drag of this pointer whose release never arrived, and the drags of pointers
        // that are gone (a finger whose release never came holds nothing any more; a
        // drag left in the list would keep the shared undo step open)
        Release(FindDrag(e.Pointer));
        foreach (var stale in drags.Where(other => other.Pointer.Captured == null).ToList())
        {
            Release(stale);
        }

        // A double click on a text edits it where it is. The label is the one the first
        // click pressed: that click selected it, and the view may have moved to bring it
        // out from under the side panel since.
        var pressedBefore = lastPressed;
        lastPressed = null;
        if (e.ClickCount == 2 && InPlaceTextEditor.CanEdit(pressedBefore))
        {
            InPlaceTextEditor.Open((LabelBase)pressedBefore, e.GetPosition(ParentCanvas));
            return;
        }

        // (not on a show/hide box, which two quick clicks tick and untick)
        if (e.ClickCount == 2 && !(Drawing.Figures.HitTest(Coordinates(e, false, false, false)) is ShowHideControl))
        {
            Drawing.ZoomToFit();
            return;
        }

        var drag = PrepareDrag(e);
        lastPressed = drag.Found;
        drags.Add(drag);
    }

    // what the last press found, for a double click
    IFigure lastPressed;

    /// <summary>What a press here would hold: the figure under it and what a drag moves (nothing is changed)</summary>
    Drag PrepareDrag(MouseButtonEventArgs e)
    {
        var drag = new Drag() { Pointer = e.Pointer };
        drag.OffsetFromFigureLeftTopCorner = Coordinates(e, false, false, false);
        drag.OldCoordinates = drag.OffsetFromFigureLeftTopCorner;
        drag.CoordinatesOnMouseDown = drag.OffsetFromFigureLeftTopCorner;

        IEnumerable<IFigure> roots = null;
        bool isLocked = false;

        var found = FindPressedFigure(
            drag.OffsetFromFigureLeftTopCorner,
            e.Pointer.Type == PointerType.Touch,
            out drag.PressedFixedLabel,
            out bool isCaption);
        drag.Found = found;
        var captions = isCaption ? new PinnedLabelScroll(Drawing) : null;

        // a press on a figure selected with others drags them all
        var selection = SelectionToDrag(drag);
        drag.DraggingSelection = selection != null;

        // a figure with parts (a slider) says which of them the press takes
        IMovable oneMovable = found is IMovableParts parts
            ? parts.FindMovablePart(drag.OffsetFromFigureLeftTopCorner)
            : found as IMovable;
        if (selection != null)
        {
            roots = FindSelectionRoots(drag, selection, out isLocked);
        }
        else if (oneMovable != null && (found.Locked || oneMovable.AllowMove()))
        {
            if (found.Locked)
            {
                isLocked = true;
            }
            else if (oneMovable.AllowMove())
            {
                if (oneMovable is IPoint)
                {
                    // when we drag a point, we want it to snap to the cursor
                    // so that the point center is directly under the tip of the mouse
                    drag.OffsetFromFigureLeftTopCorner = new Point();
                    drag.OldCoordinates = oneMovable.Coordinates;
                }
                else
                {
                    // however when we drag other stuff (such as text labels)
                    // we want the mouse to always touch the part of the draggable
                    // where it first touched during MouseDown
                    // we don't want the draggable to "snap" to the cursor like points do
                    drag.OffsetFromFigureLeftTopCorner = drag.OffsetFromFigureLeftTopCorner.Minus(oneMovable.Coordinates);
                }
                roots = DependencyAlgorithms.FindRoots(f => f.Dependents, found);

                // (a show/hide box holds figures, it isn't built on them: locked, it
                // stays where it is, and they still move)
                if (roots.All(root => !root.Locked || root is ShowHideControl))
                {
                    drag.Moving.Add(oneMovable);
                    roots = found.AsEnumerable();
                }
                else
                {
                    isLocked = true;
                }
            }
        }
        else if (found != null)
        {
            if (!found.Locked)
            {
                // a Number has no place to move; the drag goes to the points
                var allRoots = DependencyAlgorithms.FindRoots(f => f.Dependencies, found)
                    .Where(root => !(root is INumber))
                    .ToArray();

                // A point by coordinates stays where its X and Y say: the drag goes to
                // the other roots (a segment from such a point to a free one turns
                // about it), and a figure built on such points alone doesn't move at
                // all - nor does the view. (It used to "move" them: nothing happened,
                // but each drag left an undo step that undid nothing.)
                roots = allRoots.Where(root => !(root is PointByCoordinates)).ToArray();
                if (roots.IsEmpty() && !allRoots.IsEmpty())
                {
                    isLocked = true;
                }
                else if (roots.All(root => root is IMovable))
                {
                    if (roots.All(root => ((IMovable)root).AllowMove()))
                    {
                        drag.Moving.AddRange(roots.OfType<IMovable>());
                    }
                    else
                    {
                        isLocked = true;
                    }
                }
            }
            else
            {
                isLocked = true;
            }
        }

        if (roots != null)
        {
            drag.ToRecalculate = DependencyAlgorithms.FindDescendants(f => f.Dependents, roots);
            drag.ToRecalculate.Reverse();
        }
        else
        {
            drag.ToRecalculate = null;
        }

        if (drag.Moving.IsEmpty() && !isLocked && !Drawing.CoordinateGrid.Locked)
        {
            drag.Moving.Add(Drawing.CoordinateSystem);
            if (captions != null)
            {
                drag.Moving.Add(captions);
            }

            drag.ToRecalculate = null;
        }

        return drag;
    }

    /// <summary>
    /// What is under the cursor, for the status and for Tab (<see cref="ClickChoice"/>):
    /// every figure there in the order a press takes them (the one a press takes first),
    /// so that the status names what a click would select, and Tab reaches a figure
    /// under another - the segment under its end, a handle of a Bezier path on its anchor,
    /// where a click leaves it (both handles of a corner; also while the handles don't
    /// show: over an anchor whose handles are on it, a path of clicks only looks like
    /// segments). Not an axis or a caption, on which a press pans. A finger can't hover
    /// and a tap with a choice asks in a menu, so it is offered the handles only: a tap
    /// on a point on a circle is the point, as it was.
    /// </summary>
    protected override IReadOnlyList<object> FindClickOptions(MouseEventArgs e)
    {
        // (nothing while a press is held: the status says what the keys do to the dragged point)
        if (drags.Count > 0)
        {
            return new object[0];
        }

        var coordinates = Coordinates(e, false, false, false);
        var figures = Drawing.Figures.HitTestAll(coordinates, f => f.Visible && f.IsHitTestVisible && !(f is AxisLine) && !IsCaption(f));
        bool isTouch = e.Pointer.Type == PointerType.Touch;

        // a path gives one part at a place: both handles of an anchor are there; and an
        // anchor's own handles on it are there, shown or not
        var options = new List<object>();
        void Add(IFigure candidate)
        {
            if (!options.Contains(candidate))
            {
                options.Add(candidate);
            }
        }

        foreach (var figure in figures)
        {
            if (figure is BezierPath.BezierPathHandle handle)
            {
                foreach (var shown in handle.Owner.Handles.Where(h => h.Visible && h.Exists && h.HitTest(coordinates) != null))
                {
                    Add(shown);
                }

                continue;
            }

            if (isTouch && !(figure is IPoint))
            {
                continue;
            }

            Add(figure);
            if (figure is IPoint && !figure.Locked)
            {
                foreach (var path in Drawing.Figures.OfType<BezierPath>().Where(path => path.Visible && path.IsAnchor(figure)))
                {
                    foreach (var onAnchor in path.HandlesOn(figure, coordinates))
                    {
                        Add(onAnchor);
                    }
                }
            }
        }

        if (isTouch && !(options.Count > 1 && options.Any(option => option is BezierPath.BezierPathHandle)))
        {
            return new object[0];
        }

        return options;
    }

    /// <summary>A label a gallery drawing came with, pinned to the screen: paper to the Drag tool (a press on it pans)</summary>
    bool IsCaption(IFigure figure)
    {
        return figure is Label { Pin: not LabelPin.None } label && Drawing.FixedLabels.Contains(label);
    }

    /// <summary>
    /// The figures selected with the one pressed on, each a whole, which a drag moves
    /// together; null when the press is not on a selection of several
    /// </summary>
    List<IFigure> SelectionToDrag(Drag drag)
    {
        // (a handle of a Bezier path stands for its path in a selection, but a press on
        // one is for the handle)
        if (drag.Found == null || drag.Found is BezierPath.BezierPathHandle)
        {
            return null;
        }

        var clicked = FigureParts.SelectionTarget(drag.Found);
        if (!clicked.Selected && !FigureParts.Whole(clicked).Selected)
        {
            return null;
        }

        // a caption stays where it is pinned, as when the press is on it
        var selection = FigureParts.Wholes(Drawing.GetSelectedFigures())
            .Where(figure => !(figure is Label { Pin: not LabelPin.None } label && Drawing.FixedLabels.Contains(label)))
            .ToList();
        return selection.Count > 1 ? selection : null;
    }

    /// <summary>
    /// What a drag of the selection moves, into <see cref="moving"/>: the points the
    /// figures are built on, free labels, the sliders selected - each once, so that the
    /// whole selection goes by the same offset and keeps its shape. Points by coordinates
    /// stay where they are, as in the drag of one figure; one figure locked, or a point
    /// that can't move, holds them all. Returns the roots, for what to recalculate.
    /// </summary>
    IList<IFigure> FindSelectionRoots(Drag drag, List<IFigure> selection, out bool isLocked)
    {
        // a show/hide box moves itself, not the figures it holds
        var boxes = selection.OfType<ShowHideControl>().ToList<IFigure>();
        var roots = DependencyAlgorithms.FindRoots(f => f.Dependencies, selection.Except(boxes))
            .Where(root => root is IMovableParts
                ? selection.Contains(root)
                : !(root is INumber) && !(root is PointByCoordinates))
            .Concat(boxes)
            .ToList();
        var movables = roots
            .Select(root => root is IMovableParts parts ? parts.WholePart : root as IMovable)
            .Where(movable => movable != null)
            .ToList();
        isLocked = movables.IsEmpty()
            || selection.Any(figure => figure.Locked)
            || movables.Any(movable => !movable.AllowMove());
        if (!isLocked)
        {
            drag.Moving.AddRange(movables);
        }

        return roots;
    }

    public delegate void DraggerMouseMoveHandler(Point previousPoint, ref Point currentPoint);

    public event DraggerMouseMoveHandler PreviewMouseMoveCoordinates;

    public override void MouseMove(object sender, MouseEventArgs e)
    {
        var drag = FindDrag(e.Pointer);
        if (drag == null)
        {
            return;
        }

        // A move with no button down is no drag. The release of this press never
        // arrived (the tool was changed by a key and back while the button was held,
        // the window lost the mouse): the point went on following the cursor, every
        // move a recorded one, until the next click.
        if (!e.GetCurrentPoint(ParentCanvas).Properties.IsLeftButtonPressed)
        {
            Release(drag);
            return;
        }

        var moving = drag.Moving;
        var found = drag.Found;
        var currentCoordinates = Coordinates(e);

        currentCoordinates = AdjustCoordinates(drag, currentCoordinates);

        if (!drag.StartedMoving)
        {
            if (currentCoordinates == drag.CoordinatesOnMouseDown)
            {
                return;
            }

            // A press that wobbles is still a click. Without this a click with a hand
            // not perfectly still - most of them - moved what it was meant to select by
            // a pixel (a point jumped under the cursor), or panned the view by one, did
            // not select anything, and left an undo step that seemed to undo nothing.
            var coordinateSystem = Drawing.CoordinateSystem;
            var wobble = coordinateSystem.ToPhysical(Coordinates(e, false, false, false))
                .Distance(coordinateSystem.ToPhysical(drag.CoordinatesOnMouseDown));
            if (wobble < (e.Pointer.Type == PointerType.Touch ? TouchSlop : DragThreshold))
            {
                return;
            }

            drag.StartedMoving = true;
            if (!moving.IsEmpty())
            {
                // one undo step for the gesture: the moves of two fingers on one
                // figure, each pulling its roots its own way, come interleaved and
                // can't merge into one action
                dragTransaction ??= Transaction.Create(Drawing.ActionManager, false);
            }

            // (a selection goes as one piece: Alt does nothing to its points)
            if (found is IPoint && !found.Locked && !moving.IsEmpty() && !drag.DraggingSelection)
            {
                // the handles next to an anchor of a Bezier path show while it is dragged
                BezierPath.ShowHandlesWhileDragging(Drawing, found);

                // (a finger has no keys to hold)
                var hint = e.Pointer.Type == PointerType.Touch ? null : ModifierHint(found);
                if (hint != null)
                {
                    Drawing.RaiseStatusNotification(hint);
                    showsModifierHint = true;
                }
            }

            // a text: Alt attaches it to a figure, or lets an attached one go
            if (found is LabelBase && !moving.IsEmpty() && !drag.DraggingSelection && e.Pointer.Type != PointerType.Touch)
            {
                var hint = TextHint(found);
                if (hint != null)
                {
                    Drawing.RaiseStatusNotification(hint);
                    showsModifierHint = true;
                }
            }
        }
        if (!moving.IsEmpty())
        {
            var offset = currentCoordinates.Minus(drag.OldCoordinates);
            var snapping = PointToSnap(drag, currentCoordinates);

            // (a released point is dragged instead of the one pressed)
            moving = drag.Moving;
            found = drag.Found;
            if (snapping != null)
            {
                var target = drag.Snap != null ? drag.Snap.Coordinates : currentCoordinates;
                offset = target.Minus(snapping.Coordinates);
            }
            else if (moving.Count == 1 && moving[0] is PointLabel pointLabel)
            {
                // A point label is confined to an orbit around its point, so a relative
                // move would drift away from the cursor once the limit is hit. Move it
                // to where the cursor wants it (as far as allowed) and record exactly
                // that, so that undo restores the precise position.
                var desired = currentCoordinates.Minus(drag.OffsetFromFigureLeftTopCorner);
                offset = pointLabel.ClampPosition(desired).Minus(pointLabel.Coordinates);
            }
            else if (moving.Count == 1 && moving[0] is BezierPath.BezierPathHandle handle)
            {
                // the handle across the anchor follows as its mirror image, unless Alt
                // is held (a corner); with Alt it sits on a point under the cursor, and
                // dropped there is that point from then on
                handle.MirrorsOpposite = !IsAltPressed();
                drag.Snap = HandleSnap(handle, currentCoordinates);
                if (drag.Snap != null)
                {
                    offset = drag.Snap.Coordinates.Minus(handle.Coordinates);
                }
            }
            else if (moving.Count == 1 && moving[0] is PointOnFigure pointOnFigure)
            {
                // The same for a point on a figure, which stops at the end of a segment or
                // ray: record where it lands (as MoveToCore puts it), not where the cursor
                // went, or every step past the end adds to what undo moves back.
                var figure = pointOnFigure.LinearFigure;
                var landing = figure.GetPointFromParameter(
                    figure.GetNearestParameterFromPoint(pointOnFigure.Coordinates.Plus(offset)));
                offset = landing.Minus(pointOnFigure.Coordinates);
            }
            else if (moving.Count == 1 && moving[0] is PointInRegion inside && inside.Region != null)
            {
                // and for a point kept inside a shape, which stops at its outline
                var landing = inside.Region.NearestPointInside(inside.Coordinates.Plus(offset));
                if (landing.Exists())
                {
                    offset = landing.Minus(inside.Coordinates);
                }
            }

            AimText(drag, Coordinates(e, false, false, false));
            Actions.Move(Drawing, moving, offset, drag.ToRecalculate);
        }

        // OK attention here. This is a very tricky spot. At the beginning
        // of this method, we call Coordinates(e) to get the logical mouse
        // coordinates. We could just reuse currentCoordinates, BUT!
        // If you're dragging the coordinate plane itself, the Origin changes
        // so you'll have to re-get the point coordinates in the new
        // coordinate system.
        drag.OldCoordinates = Coordinates(e);
        if (!drag.DraggingSelection
            && moving.Count == 1
            && moving[0] is IPoint
            && found != null
            && (found == moving[0] || found is IMovableParts))
        {
            drag.OldCoordinates = moving[0].Coordinates;
        }
    }

    private Point AdjustCoordinates(Drag drag, Point currentCoordinates)
    {
        if (PreviewMouseMoveCoordinates != null)
        {
            PreviewMouseMoveCoordinates(drag.OldCoordinates, ref currentCoordinates);
        }
        return currentCoordinates;
    }

    public override void MouseUp(object sender, MouseButtonEventArgs e)
    {
        var drag = FindDrag(e.Pointer);
        if (drag == null)
        {
            return;
        }

        var found = drag.Found;
        try
        {
            if (drag.Snap != null && IsAltPressed() && found is FreePoint dragged)
            {
                PointSnapping.Snap(dragged, drag.Snap);
            }
            else if (drag.Snap?.ExistingPoint != null && IsAltPressed() && found is BezierPath.BezierPathHandle handle && dragTransaction != null)
            {
                handle.Owner.UsePointAsHandle(handle, drag.Snap.ExistingPoint);
            }
            else if (IsAltPressed() && found is LabelBase text && dragTransaction != null)
            {
                // (in the drag's step: one undo puts the text back as it was, where it was)
                if (drag.AttachTarget != null)
                {
                    TextAttachment.Attach(text, drag.AttachTarget);
                }
                else if (drag.DetachesText && text is AttachedLabel attached)
                {
                    TextAttachment.Detach(attached);
                }
            }
        }
        catch
        {
            // (the drag is over also when the drop threw, or the next move would go on with it)
            Release(drag);
            throw;
        }
        finally
        {
            // also when the drop threw: left open, the drag's transaction would take in
            // everything done from then on, and nothing of it could be undone
            EndDrag(drag);
        }

        // a press that didn't become a drag; on a show/hide box it ticks the box, with
        // Ctrl it selects it as any figure
        if (!drag.StartedMoving)
        {
            if ((found ?? drag.PressedFixedLabel) is ShowHideControl box && !IsCtrlPressed())
            {
                box.Click();
            }
            else
            {
                UpdateSelection(drag);
                Drawing.RaiseSelectionChanged(Drawing.GetSelectedFigures());
            }
        }

        Release(drag);
    }

    public override void Stopping()
    {
        // (a text being edited goes in: Escape alone leaves it out)
        InPlaceTextEditor.Current?.Commit();
        ReleaseAll();
        AbandonGesture();
        base.Stopping();
    }

    #region Lasso: a drag with the right button

    // a lasso being drawn, null otherwise
    RegionSelection lasso;

    protected override bool MouseRightDragStarted(object sender, PointerEventArgs press)
    {
        // (not while a press of the left button drags something)
        if (drags.Count > 0)
        {
            return false;
        }

        lasso = new RegionSelection(
            Drawing,
            SelectionRegion.Lasso(press.GetPosition(ParentCanvas)),
            showsKeyHint: true);
        return true;
    }

    protected override void MouseRightDragMoved(object sender, PointerEventArgs e)
    {
        lasso?.Extend(e.GetPosition(ParentCanvas), removing: IsAltPressed());
    }

    protected override void MouseRightDragEnded(object sender, PointerEventArgs e)
    {
        if (lasso == null)
        {
            return;
        }

        var finished = lasso;
        lasso = null;
        finished.Extend(e.GetPosition(ParentCanvas), removing: IsAltPressed());
        finished.Finish();
    }

    public override bool AbandonGesture()
    {
        // Escape in a text edited where it is leaves the label as it was
        if (InPlaceTextEditor.Current != null)
        {
            InPlaceTextEditor.Current.Cancel();
            return true;
        }

        if (lasso == null)
        {
            return false;
        }

        lasso.Cancel();
        lasso = null;
        return true;
    }

    #endregion

    /// <summary>The tool on another drawing (the next of the gallery): what it held of the old one is let go</summary>
    public override Drawing Drawing
    {
        get => base.Drawing;
        set
        {
            if (value != base.Drawing)
            {
                ReleaseAll();
                AbandonGesture();
            }

            base.Drawing = value;
        }
    }

    /// <summary>The press is over: nothing is held by it any more</summary>
    void Release(Drag drag)
    {
        if (drag == null)
        {
            return;
        }

        EndDrag(drag);
        drags.Remove(drag);

        // (the handles of the point another finger still drags stay)
        BezierPath.ShowHandlesWhileDragging(Drawing, drags.LastOrDefault(other => other.StartedMoving)?.Found);
    }

    void ReleaseAll()
    {
        foreach (var drag in drags.ToList())
        {
            Release(drag);
        }
    }

    /// <summary>
    /// The drag is over: the last one of the fingers down together ends the undo step
    /// they share (also when the drop threw, and the drag has left the list already)
    /// </summary>
    void EndDrag(Drag drag)
    {
        drag.Snap = null;
        drag.AttachTarget = null;
        drag.DetachesText = false;

        // (a drag whose pointer is gone without a release holds nothing: it must not
        // keep the step open)
        drags.RemoveAll(other => other != drag && other.Pointer.Captured == null);
        if (drags.Any(other => other != drag))
        {
            return;
        }

        if (dragTransaction != null)
        {
            dragTransaction.Commit();
            dragTransaction = null;
        }

        if (showsModifierHint)
        {
            showsModifierHint = false;
            Drawing?.RaiseStatusNotification(HintText);
        }
    }

    // the status says what Shift and Alt do to the dragged point, until the drop
    bool showsModifierHint;

    /// <summary>
    /// What holding a key does to the point that is dragged, for the status: Shift puts a
    /// free point on the grid, Alt snaps it to a figure or lets go of the figures it is
    /// on (<see cref="PointToSnap"/>); null for a point the keys do nothing to
    /// </summary>
    static string ModifierHint(IFigure figure)
    {
        if (figure is BezierPath.BezierPathHandle)
        {
            return "Hold Shift to snap to grid. Hold " + KeyNames.Alt + " to leave the handle across the anchor where it is, or to drop this one on a point.";
        }

        if (!(figure is PointBase point))
        {
            return null;
        }

        // (a point a locus is drawn from stays what it is: Alt does nothing to it)
        bool altApplies = !PointSnapping.IsHeldByLocus(point);
        if (PointSnapping.CanRelease(point))
        {
            return altApplies ? "Hold " + KeyNames.Alt + " to detach the point." : null;
        }

        // (not a point on a figure whose release was refused: the keys do nothing to it)
        if (PointSnapping.IsFreePoint(point))
        {
            return altApplies
                ? "Hold Shift to snap to grid. Hold " + KeyNames.Alt + " to snap to a figure."
                : "Hold Shift to snap to grid.";
        }

        return null;
    }

    #region Alt-drag: snapping and releasing points

    /// <summary>
    /// With Alt held, a dragged point is a free point that snaps (<see cref="PointSnapping"/>):
    /// one tied to figures is released first, and <see cref="snap"/> says where it goes -
    /// into another point, onto a figure, an intersection or the middle of a segment,
    /// whatever the Point tool would make there. Returns the free point, or null when Alt
    /// doesn't apply (not held, not a point, a locked one).
    /// </summary>
    FreePoint PointToSnap(Drag drag, Point cursor)
    {
        var previous = drag.Snap;
        drag.Snap = null;
        // (a point a locus is drawn from stays what it is: Alt does nothing to it; nor to
        // a point dragged with the rest of a selection, which goes as one piece - Alt
        // released the point pressed, which then moved alone, off what it was on)
        if (!IsAltPressed()
            || drag.DraggingSelection
            || !(drag.Found is PointBase point)
            || drag.Found.Locked
            || dragTransaction == null
            || PointSnapping.IsHeldByLocus(point))
        {
            return null;
        }

        FreePoint free;
        if (PointSnapping.CanRelease(point))
        {
            free = PointSnapping.Release(point);
            DragInstead(drag, free);
        }
        // a free point proper: a point on a figure is a FreePoint too, and one whose release
        // was refused (an expression reads its Parameter) was snapped into an intersection
        else if (PointSnapping.IsFreePoint(point) && drag.Moving.Count == 1 && drag.Moving[0] == point)
        {
            free = (FreePoint)point;
        }
        else
        {
            return null;
        }

        drag.Snap = PointPlacement.FindSnap(
            Drawing,
            cursor,
            figure => !figure.DependsOn(free),
            target => PointSnapping.CanJoin(free, target));
        if (drag.Snap == null && previous != null)
        {
            drag.Snap = Stick(previous, cursor);
        }

        return free;
    }

    /// <summary>
    /// With Alt, a dragged handle of a Bezier path sits on a point under the cursor
    /// (<see cref="BezierPath.CanUseAsHandle"/>), the one a drop makes it; null when there is
    /// none, or Alt is not held
    /// </summary>
    PointPlacement HandleSnap(BezierPath.BezierPathHandle handle, Point cursor)
    {
        if (!IsAltPressed() || dragTransaction == null)
        {
            return null;
        }

        var target = Drawing.Figures.HitTestMany(cursor)
            .OfType<IPoint>()
            .FirstOrDefault(point => point.IsHitTestVisible && handle.Owner.CanUseAsHandle(handle, point));
        return target != null ? PointPlacement.Existing(target) : null;
    }

    /// <summary>The point that replaced the one pressed on is what the drag moves now</summary>
    void DragInstead(Drag drag, PointBase point)
    {
        drag.Found = point;
        drag.Moving = new List<IMovable>() { (IMovable)point };
        drag.ToRecalculate = DependencyAlgorithms.FindDescendants(f => f.Dependents, point.AsEnumerable<IFigure>());
        drag.ToRecalculate.Reverse();
    }

    /// <summary>
    /// A snapped point lets go only at <see cref="StickyReach"/> times the reach that took
    /// it, so that it doesn't flicker on and off at the edge
    /// </summary>
    PointPlacement Stick(PointPlacement previous, Point cursor)
    {
        // an intersection or a midpoint stays put: nothing it is made of is built on the dragged point
        var placement = previous.Kind == PointPlacementKind.OnFigure
            ? PointPlacement.OnFigure((ILinearFigure)previous.Sources[0], cursor)
            : previous;
        var reach = StickyReach * Drawing.CoordinateSystem.CursorTolerance;
        return placement.Kind != PointPlacementKind.Free && placement.Coordinates.Distance(cursor) <= reach
            ? placement
            : null;
    }

    /// <summary>What holding Alt does to a dragged text, for the status; null for a text it does nothing to (a caption pinned to the screen, one with [...] parts)</summary>
    static string TextHint(IFigure text)
    {
        if (!TextAttachment.CanAttach(text))
        {
            return null;
        }

        return text is AttachedLabel
            ? "Hold " + KeyNames.Alt + " to attach the text to another figure, or to let go of this one over the paper."
            : "Hold " + KeyNames.Alt + " to attach the text to a figure.";
    }

    /// <summary>
    /// With Alt held, a dragged text goes onto the figure under the cursor
    /// (<see cref="TextAttachment"/>), or an attached one lets go of its figure over the
    /// paper: the halo shows the figure, the status says what a drop does. Nothing changes
    /// before the drop, as for a point dragged with Alt.
    /// </summary>
    void AimText(Drag drag, Point cursor)
    {
        var target = drag.AttachTarget;
        bool detaches = drag.DetachesText;
        drag.AttachTarget = null;
        drag.DetachesText = false;
        if (drag.DraggingSelection
            || dragTransaction == null
            || drag.Moving.Count != 1
            || drag.Found is not LabelBase text)
        {
            return;
        }

        // (a free text with [...] parts says why Alt does nothing to it, while Alt is held)
        if (!TextAttachment.CanAttach(text))
        {
            if (IsAltPressed() && TextAttachment.WhyNot(text) is string why)
            {
                showsModifierHint = true;
                Drawing.RaiseStatusNotification(why);
            }

            return;
        }

        if (IsAltPressed())
        {
            var under = Drawing.Figures
                .HitTestAll(cursor, figure => figure != text && figure.Visible && figure.IsHitTestVisible)
                .Select(figure => Drawing.Figures.FindTopLevel(figure))
                .FirstOrDefault(TextAttachment.CanAttachTo);
            var own = (text as AttachedLabel)?.Figure;
            if (under != null && under != own)
            {
                drag.AttachTarget = under;
            }
            else if (under == null && own != null)
            {
                drag.DetachesText = true;
            }
        }

        if (drag.AttachTarget == target && drag.DetachesText == detaches)
        {
            return;
        }

        showsModifierHint = true;
        Drawing.RaiseStatusNotification(drag.AttachTarget != null
            ? "Release to attach the text to " + ConstructionText.Of(drag.AttachTarget) + "."
            : drag.DetachesText
                ? "Release to let go of " + ConstructionText.Of(((AttachedLabel)text).Figure) + "."
                : TextHint(text));
    }

    /// <summary>The halo and the faint point of the snap, while dragging with Alt</summary>
    protected override PointPlacement GetClickPreview(MouseEventArgs e)
    {
        return FindDrag(e.Pointer)?.Snap;
    }

    /// <summary>
    /// Over a point to join, the dragged point sits right on it and the halo goes on that
    /// point; nothing is joined before the drop. Or the figure chosen with Tab under
    /// another, which a press takes instead of the topmost: the halo shows which. Over an
    /// angle's number, its mark: which angle the number says, where numbers crowd.
    /// </summary>
    protected override IFigure GetFigureToPick(MouseEventArgs e)
    {
        var drag = FindDrag(e.Pointer);
        var picked = drag?.Snap?.ExistingPoint ?? drag?.AttachTarget ?? (Choice.Index > 0 ? Choice.Current as IFigure : null);
        if (picked != null || drag != null)
        {
            return picked;
        }

        return Choice.Current is AngleMeasurement number ? AngleArc.FindCompanion(number) : null;
    }

    #endregion

    /// <summary>
    /// The menu of the figure under the cursor, or of the paper, like the point and figure
    /// menus of the original DG (<see cref="ContextMenus"/>)
    /// </summary>
    public override void MouseRightClick(object sender, MouseButtonEventArgs e)
    {
        var coordinates = Coordinates(e, false, false, false);
        var hit = Drawing.Figures.HitTest(coordinates);

        // (the figure chosen with Tab, as a click takes it)
        if (Choice.Index > 0 && Choice.Current is IFigure chosen)
        {
            hit = chosen;
        }

        if (hit is AxisLine)
        {
            hit = null;
        }

        if (hit == null)
        {
            ContextMenus.ForPaper(Drawing).Open(ParentCanvas);
            return;
        }

        // a side of a Bezier path is selected by itself, as a click selects it, but what
        // the menu does (hide, lock, delete) is done to the path
        var clicked = FigureParts.SelectionTarget(hit);
        if (!clicked.Selected)
        {
            Drawing.Figures.ClearSelection();
            clicked.Selected = true;
            Drawing.RaiseSelectionChanged(Drawing.GetSelectedFigures());
        }

        var choices = Drawing.Figures
            .HitTestAll(coordinates, f => f.Visible && f.IsHitTestVisible && !(f is AxisLine))
            .Select(FigureParts.SelectionTarget)
            .Distinct()
            .ToList();
        ContextMenus.ForFigure(Drawing, FigureParts.Whole(clicked), choices).Open(ParentCanvas);
    }

    public override void KeyDown(object sender, KeyEventArgs e)
    {
        var selectedFigures = Drawing.GetSelectedFigures();
        if (e.Key == Key.Delete && !selectedFigures.IsEmpty())
        {
            Drawing.DeleteSelection();
            e.Handled = true;
        }
    }

    private void UpdateSelection(Drag drag)
    {
        // a click on a handle of a Bezier path leaves the selection as it is: the handle
        // shows while its anchor is selected, and would go with the anchor
        if (drag.Found is BezierPath.BezierPathHandle)
        {
            return;
        }

        // a vertex or a side of a regular polygon by itself, the inside for the polygon
        SelectClicked(Drawing, drag.Found != null ? FigureParts.SelectionTarget(drag.Found) : drag.PressedFixedLabel);
    }

    /// <summary>
    /// What a click selects (null when it is on nothing): that alone, or with Ctrl it goes
    /// into the selection or out of it. Also the click of the selection tools
    /// (<see cref="RegionSelector"/>).
    /// </summary>
    public static void SelectClicked(Drawing drawing, IFigure clicked)
    {
        if (IsCtrlPressed())
        {
            // a part of a polygon selected whole: the polygon goes out of the selection
            var whole = clicked != null ? FigureParts.Whole(clicked) : null;
            if (whole != clicked && whole.Selected)
            {
                whole.Selected = false;
            }
            else if (clicked != null)
            {
                clicked.Selected = !clicked.Selected;
            }
        }
        else
        {
            drawing.Figures.ClearSelection();
            if (clicked != null)
            {
                clicked.Selected = true;
            }
        }
    }

    public override FrameworkElement CreateIcon()
    {
        Point[] points = 
        [
            new Point(10, 5),
            new Point(10, 21),
            new Point(14, 17),
            new Point(18, 25),
            new Point(19, 25),
            new Point(20, 24),
            new Point(17, 17),
            new Point(17, 16),
            new Point(21, 16),
            new Point(10, 5)
        ];
        var builder = IconBuilder.BuildIcon();
        var polygon = builder.AddPolygon(
                points.Select(p => new Point(p.X / 32, p.Y / 32)));
        polygon.Fill = new SolidColorBrush(Colors.White);
        polygon.Stroke = new SolidColorBrush(Color.FromRgb(0x30, 0x36, 0x40));
        polygon.StrokeThickness = 1.5;
        polygon.StrokeJoin = PenLineJoin.Round;
        return builder.Canvas;
    }

    public override string Name => "Drag";

    public override string HintText => "Use this tool to drag points and figures. Drag with the right button to select figures with a lasso.";
}