using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Input;
using Avalonia.Media;

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

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
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

            IMovable oneMovable = found as IMovable;
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
                    roots = DependencyAlgorithms.FindRoots(f => f.Dependencies, found)
                        .Where(root => !(root is INumber))
                        .ToArray();
                    if (roots.All(root => root is IMovable))
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
                startedMoving = true;
            }
            if (!moving.IsEmpty())
            {
                var offset = currentCoordinates.Minus(oldCoordinates);
                if (moving.Count == 1 && moving[0] is PointLabel pointLabel)
                {
                    // A point label is confined to an orbit around its point, so a relative
                    // move would drift away from the cursor once the limit is hit. Move it
                    // to where the cursor wants it (as far as allowed) and record exactly
                    // that, so that undo restores the precise position.
                    var desired = currentCoordinates.Minus(offsetFromFigureLeftTopCorner);
                    offset = pointLabel.ClampPosition(desired).Minus(pointLabel.Coordinates);
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
                && found == moving[0])
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
            if (Coordinates(e) == coordinatesOnMouseDown)
            {
                UpdateSelection();
                Drawing.RaiseSelectionChanged(Drawing.GetSelectedFigures());
            }

            startedMoving = false;
            moving = null;
            found = null;
        }

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