using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Input;
using Avalonia.Media;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Selection)]
    [Order(1)]
    public partial class Dragger : Behavior
    {
        protected List<IMovable> moving = null;
        IFigure found = null;
        List<IFigure> toRecalculate = null;
        Point offsetFromFigureLeftTopCorner;
        protected Point oldCoordinates;
        protected Point coordinatesOnMouseDown;
        bool startedMoving = false;

        // A point drag is one undo step, with the releases and the snap of an Alt-drag in it
        Transaction dragTransaction;

        // Where the point dragged with Alt goes if dropped now; null while it is free
        PointPlacement snap;

        /// <summary>In cursor tolerances: how far a snapped point may be pulled before it lets go</summary>
        public static double StickyReach = 2;

        /// <summary>In pixels: how far the cursor goes from where it was pressed before the press is a drag and not a click</summary>
        public static double DragThreshold = 3;

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            // a drag whose release never arrived
            EndDrag();

#if !SILVERLIGHT
            if (e.ClickCount == 2)
            {
                Drawing.CoordinateSystem.ZoomExtend();
                return;
            }
#endif
            offsetFromFigureLeftTopCorner = Coordinates(e, false, false, false);
            oldCoordinates = offsetFromFigureLeftTopCorner;
            coordinatesOnMouseDown = offsetFromFigureLeftTopCorner;
            startedMoving = false;

            moving = new List<IMovable>();
            IEnumerable<IFigure> roots = null;
            bool isLocked = false;

            found = Drawing.Figures.HitTest(offsetFromFigureLeftTopCorner);

            // labels that can't be dragged are paper: the drag moves the view, and one that
            // starts on a caption takes the captions along (Drawing.FixedLabels)
            PinnedLabelScroll captions = null;
            if (Drawing.FixedLabels && found is LabelBase)
            {
                if (found is Label { Pin: not LabelPin.None })
                {
                    captions = new PinnedLabelScroll(Drawing);
                }

                found = null;
            }

            // a figure with parts (a slider) says which of them the press takes
            IMovable oneMovable = found is IMovableParts parts
                ? parts.FindMovablePart(offsetFromFigureLeftTopCorner)
                : found as IMovable;
            if (oneMovable != null && (found.Locked || oneMovable.AllowMove()))
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
                        offsetFromFigureLeftTopCorner = new Point();
                        oldCoordinates = oneMovable.Coordinates;
                    }
                    else
                    {
                        // however when we drag other stuff (such as text labels)
                        // we want the mouse to always touch the part of the draggable
                        // where it first touched during MouseDown
                        // we don't want the draggable to "snap" to the cursor like points do
                        offsetFromFigureLeftTopCorner = offsetFromFigureLeftTopCorner.Minus(oneMovable.Coordinates);
                    }
                    roots = DependencyAlgorithms.FindRoots(f => f.Dependents, found);
                    if (roots.All(root => (!root.Locked)))
                    {
                        moving.Add(oneMovable);
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
                            moving.AddRange(roots.OfType<IMovable>());
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
                toRecalculate = DependencyAlgorithms.FindDescendants(f => f.Dependents, roots);
                toRecalculate.Reverse();
            }
            else
            {
                toRecalculate = null;
            }

            if (moving.IsEmpty() && !isLocked && !Drawing.CoordinateGrid.Locked)
            {
                moving.Add(Drawing.CoordinateSystem);
                if (captions != null)
                {
                    moving.Add(captions);
                }

                //var allFigures = Drawing.Figures.GetAllFiguresRecursive();
                //roots = DependencyAlgorithms.FindRoots(f => f.Dependencies, allFigures);
                //moving.AddRange(roots.OfType<IMovable>());
                //roots = null;
                toRecalculate = null; // Figures;
            }
        }

        public delegate void DraggerMouseMoveHandler(Point previousPoint, ref Point currentPoint);

        public event DraggerMouseMoveHandler PreviewMouseMoveCoordinates;

        public override void MouseMove(object sender, MouseEventArgs e)
        {
            var currentCoordinates = Coordinates(e);

            currentCoordinates = AdjustCoordinates(currentCoordinates);

            if (!startedMoving)
            {
                if (currentCoordinates == coordinatesOnMouseDown)
                {
                    return;
                }

                // A press that wobbles is still a click. Without this a click with a hand
                // not perfectly still - most of them - moved what it was meant to select by
                // a pixel (a point jumped under the cursor), or panned the view by one, did
                // not select anything, and left an undo step that seemed to undo nothing.
                var coordinateSystem = Drawing.CoordinateSystem;
                var wobble = coordinateSystem.ToPhysical(Coordinates(e, false, false, false))
                    .Distance(coordinateSystem.ToPhysical(coordinatesOnMouseDown));
                if (wobble < DragThreshold)
                {
                    return;
                }

                startedMoving = true;
                if (found is IPoint && !found.Locked && !moving.IsEmpty())
                {
                    dragTransaction = Transaction.Create(Drawing.ActionManager, false);
                }
            }
            if (!moving.IsEmpty())
            {
                var offset = currentCoordinates.Minus(oldCoordinates);
                var snapping = PointToSnap(currentCoordinates);
                if (snapping != null)
                {
                    var target = snap != null ? snap.Coordinates : currentCoordinates;
                    offset = target.Minus(snapping.Coordinates);
                }
                else if (moving.Count == 1 && moving[0] is PointLabel pointLabel)
                {
                    // A point label is confined to an orbit around its point, so a relative
                    // move would drift away from the cursor once the limit is hit. Move it
                    // to where the cursor wants it (as far as allowed) and record exactly
                    // that, so that undo restores the precise position.
                    var desired = currentCoordinates.Minus(offsetFromFigureLeftTopCorner);
                    offset = pointLabel.ClampPosition(desired).Minus(pointLabel.Coordinates);
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

                Actions.Move(Drawing, moving, offset, toRecalculate);
            }

            // OK attention here. This is a very tricky spot. At the beginning
            // of this method, we call Coordinates(e) to get the logical mouse
            // coordinates. We could just reuse currentCoordinates, BUT!
            // If you're dragging the coordinate plane itself, the Origin changes
            // so you'll have to re-get the point coordinates in the new 
            // coordinate system.
            oldCoordinates = Coordinates(e);
            if (moving != null
                && moving.Count == 1
                && moving[0] is IPoint
                && found != null
                && (found == moving[0] || found is IMovableParts))
            {
                oldCoordinates = moving[0].Coordinates;
            }
        }

        private Point AdjustCoordinates(Point currentCoordinates)
        {
            if (PreviewMouseMoveCoordinates != null)
            {
                PreviewMouseMoveCoordinates(oldCoordinates, ref currentCoordinates);
            }
            return currentCoordinates;
        }

        public override void MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (snap != null && IsAltPressed() && found is FreePoint dragged)
                {
                    PointSnapping.Snap(dragged, snap);
                }
            }
            finally
            {
                // also when the drop threw: left open, the drag's transaction would take in
                // everything done from then on, and nothing of it could be undone
                EndDrag();
            }

            // a press that didn't become a drag (moving is null when the press was elsewhere:
            // on the ribbon, in another tool)
            if (moving != null && !startedMoving)
            {
                UpdateSelection();
                Drawing.RaiseSelectionChanged(Drawing.GetSelectedFigures());
            }

            startedMoving = false;
            moving = null;
            found = null;
        }

        public override void Stopping()
        {
            EndDrag();
            base.Stopping();
        }

        void EndDrag()
        {
            snap = null;
            if (dragTransaction != null)
            {
                dragTransaction.Commit();
                dragTransaction = null;
            }
        }

        #region Alt-drag: snapping and releasing points

        /// <summary>
        /// With Alt held, a dragged point is a free point that snaps (<see cref="PointSnapping"/>):
        /// one tied to figures is released first, and <see cref="snap"/> says where it goes -
        /// into another point, onto a figure, an intersection or the middle of a segment,
        /// whatever the Point tool would make there. Returns the free point, or null when Alt
        /// doesn't apply (not held, not a point, a locked one).
        /// </summary>
        FreePoint PointToSnap(Point cursor)
        {
            var previous = snap;
            snap = null;
            // (a point a locus is drawn from stays what it is: Alt does nothing to it)
            if (!IsAltPressed()
                || !(found is PointBase point)
                || found.Locked
                || dragTransaction == null
                || PointSnapping.IsHeldByLocus(point))
            {
                return null;
            }

            FreePoint free;
            if (PointSnapping.CanRelease(point))
            {
                free = PointSnapping.Release(point);
                DragInstead(free);
            }
            else if (point is FreePoint freePoint && moving.Count == 1 && moving[0] == point)
            {
                free = freePoint;
            }
            else
            {
                return null;
            }

            snap = PointPlacement.FindSnap(
                Drawing,
                cursor,
                figure => !figure.DependsOn(free),
                target => PointSnapping.CanJoin(free, target));
            if (snap == null && previous != null)
            {
                snap = Stick(previous, cursor);
            }

            return free;
        }

        /// <summary>The point that replaced the one pressed on is what the drag moves now</summary>
        void DragInstead(PointBase point)
        {
            found = point;
            moving = new List<IMovable>() { (IMovable)point };
            toRecalculate = DependencyAlgorithms.FindDescendants(f => f.Dependents, point.AsEnumerable<IFigure>());
            toRecalculate.Reverse();
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

        /// <summary>The halo and the faint point of the snap, while dragging with Alt</summary>
        protected override PointPlacement GetClickPreview(MouseEventArgs e)
        {
            return snap;
        }

        /// <summary>
        /// Over a point to join, the dragged point sits right on it and the halo goes on that
        /// point; nothing is joined before the drop
        /// </summary>
        protected override IFigure GetFigureToPick(MouseEventArgs e)
        {
            return snap?.ExistingPoint;
        }

        #endregion

#if !PLAYER

        /// <summary>
        /// Context menu for the figure under the cursor, like the point and figure menus
        /// of the original DG.
        /// </summary>
        public override void MouseRightClick(object sender, MouseButtonEventArgs e)
        {
            var figure = Drawing.Figures.HitTest(Coordinates(e, false, false, false));
            var menu = new Avalonia.Controls.ContextMenu();

            void Add(string header, System.Action action, bool? isChecked = null)
            {
                var item = new Avalonia.Controls.MenuItem() { Header = header };
                if (isChecked != null)
                {
                    item.ToggleType = Avalonia.Controls.MenuItemToggleType.CheckBox;
                    item.IsChecked = isChecked.Value;
                }

                item.Click += (s, args) => action();
                menu.Items.Add(item);
            }

            void Set(object target, string property, object value)
            {
                Actions.SetProperty(Drawing.ActionManager, new PropertyValue(property, target), value);
            }

            if (figure == null)
            {
                Add("Zoom to fit", () => Drawing.CoordinateSystem.ZoomExtend());
                Add("Select all", () =>
                {
                    Drawing.SelectAll();
                    Drawing.RaiseSelectionChanged(Drawing.GetSelectedFigures());
                });
            }
            else
            {
                if (!figure.Selected)
                {
                    Drawing.Figures.ClearSelection();
                    figure.Selected = true;
                    Drawing.RaiseSelectionChanged(Drawing.GetSelectedFigures());
                }

                var point = figure as PointBase ?? (figure as PointLabel)?.Dependencies.FirstOrDefault() as PointBase;
                if (point != null)
                {
                    Add("Show name", () => Set(point, "ShowName", !point.ShowName), point.ShowName);
                    Add("Show coordinates", () => Set(point, "ShowCoordinates", !point.ShowCoordinates), point.ShowCoordinates);
                    AddSnapItems(menu, point, Add);
                    menu.Items.Add(new Avalonia.Controls.Separator());
                }

                Add("Hide", () =>
                {
                    Set(figure, "Visible", false);
                    figure.Selected = false;
                    Drawing.RaiseSelectionChanged(Drawing.GetSelectedFigures());
                });
                Add(figure.Locked ? "Unlock" : "Lock", () => Set(figure, "Locked", !figure.Locked));
                if (!(figure is PointLabel))
                {
                    menu.Items.Add(new Avalonia.Controls.Separator());
                    Add("Delete", () => Drawing.DeleteSelection());
                }
            }

            menu.Open(ParentCanvas);
        }

        /// <summary>
        /// As in the original DG: "Snap to" the figure through the point (a submenu when there
        /// are several), "Convert to point by coordinates" for a free point, and "Free point" for a point
        /// tied to figures or to typed coordinates
        /// </summary>
        static void AddSnapItems(
            Avalonia.Controls.ContextMenu menu,
            PointBase point,
            System.Action<string, System.Action, bool?> add)
        {
            var figures = point is FreePoint ? PointSnapping.FiguresToSnapTo(point) : new IFigure[0];
            if (figures.Count == 1)
            {
                add("Snap to " + PointSnapping.Describe(figures[0]), () => PointSnapping.SnapTo(point, figures[0]), null);
            }
            else if (figures.Count > 1)
            {
                var snapTo = new Avalonia.Controls.MenuItem() { Header = "Snap to figure" };
                foreach (var figure in figures)
                {
                    var item = new Avalonia.Controls.MenuItem() { Header = PointSnapping.Describe(figure) };
                    item.Click += (s, args) => PointSnapping.SnapTo(point, figure);
                    snapTo.Items.Add(item);
                }

                menu.Items.Add(snapTo);
            }

            if (PointSnapping.CanConvertToPointByCoordinates(point))
            {
                add("Convert to point by coordinates", () => PointSnapping.ConvertToPointByCoordinates((FreePoint)point), null);
            }

            if (PointSnapping.CanFree(point))
            {
                add("Free point", () => PointSnapping.Release(point), null);
            }
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

#endif

        private void UpdateSelection()
        {
            if (IsCtrlPressed())
            {
                if (found != null)
                {
                    found.Selected = !found.Selected;
                }
            }
            else
            {
                Drawing.Figures.ClearSelection();
                if (found != null)
                {
                    found.Selected = true;
                }
            }
        }

        public override FrameworkElement CreateIcon()
        {
            Point[] points = 
            {
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
            };
            var builder = IconBuilder.BuildIcon();
            var polygon = builder.AddPolygon(
                    points.Select(p => new Point(p.X / 32, p.Y / 32)));
            polygon.Fill = new SolidColorBrush(Colors.White);
            polygon.Stroke = new SolidColorBrush(Color.FromRgb(0x30, 0x36, 0x40));
            polygon.StrokeThickness = 1.5;
            polygon.StrokeJoin = PenLineJoin.Round;
            return builder.Canvas;
        }

        public override string Name
        {
            get { return "Drag"; }
        }

        public override string HintText
        {
            get { return "Use this tool to drag points and figures."; }
        }
    }
}