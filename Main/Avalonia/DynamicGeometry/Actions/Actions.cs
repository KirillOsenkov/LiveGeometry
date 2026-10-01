using GuiLabs.Undo;
using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace DynamicGeometry
{
    /// <summary>
    /// This class encapsulates undoable actions
    /// Calling every method here guarantees that the method will update the Undo buffer
    /// and that the action will be undoable
    /// </summary>
    public class Actions
    {
        public static void Add(Drawing drawing, IFigure newFigure)
        {
            AddFigureAction action = new AddFigureAction(drawing, newFigure);
            drawing.ActionManager.RecordAction(action);
        }

        public static void AddMany(Drawing drawing, IEnumerable<IFigure> figures)
        {
            using (drawing.ActionManager.CreateTransaction())
            {
                foreach (var figure in figures)
                {
                    Add(drawing, figure);
                }
            }
        }

        public static void Remove(IFigure figure)
        {
            var drawing = figure.Drawing;
            RemoveFigureAction action = new RemoveFigureAction(drawing, figure);
            drawing.ActionManager.RecordAction(action);
        }

        public static void ReplaceDependency(IFigure figure, IFigure oldDependency, IFigure newDependency)
        {
            CallMethodAction action = new CallMethodAction(
                () => figure.ReplaceDependency(oldDependency, newDependency),
                () => figure.ReplaceDependency(newDependency, oldDependency));
            figure.Drawing.ActionManager.RecordAction(action);
        }

        public static void ReplaceWithExisting(IFigure existingFigure, IFigure newFigure)
        {
            Drawing drawing = existingFigure.Drawing;
            ReplaceFigureAction action = new ReplaceFigureAction(drawing, existingFigure, newFigure);
            drawing.ActionManager.RecordAction(action);
        }

        public static void ReplaceWithNew(IFigure existingFigure, IFigure newFigure)
        {
            Drawing drawing = existingFigure.Drawing;
            using (drawing.ActionManager.CreateTransaction())
            {
                Actions.Add(drawing, newFigure);
                Actions.ReplaceWithExisting(existingFigure, newFigure);
                Actions.Remove(existingFigure);
                if (newFigure is PointBase && existingFigure is PointBase || !existingFigure.HasDefaultName)
                {
                    Actions.SetProperty(drawing.ActionManager, new PropertyValue("Name", newFigure), existingFigure.Name);
                }
                else if (newFigure.HasDefaultName)
                {
                    // segment AB converted to a line is line AB: AB2 while both were there
                    Actions.SetProperty(drawing.ActionManager, new PropertyValue("Name", newFigure), newFigure.GenerateFigureName());
                }
            }
        }

        /// <summary>
        /// Puts <paramref name="replacement"/> in the drawing in place of <paramref name="point"/>:
        /// same name, same dependents - its name label included, which
        /// <see cref="ReplaceFigureAction"/> leaves alone - in one undo step. What the old point
        /// was built on goes with it when nothing else uses it (an auxiliary Number). The
        /// replacement should already be where the point is, so that nothing moves. It keeps
        /// the point's visibility, lock and a style the user chose; a default style follows the
        /// replacement's kind.
        /// </summary>
        public static void ReplacePoint(PointBase point, PointBase replacement)
        {
            var drawing = point.Drawing;

            // an expression that names the point ([A.X] in a label) holds the figure itself
            // once compiled. Its text is left alone while the names change hands (the
            // replacement has a temporary one in between), and it is compiled again when the
            // replacement has the name, or on undo when the point is back: the first thing
            // and the last either way, hence the two actions
            var expressions = point.Dependents.OfType<IRenamableExpressions>().ToArray();
            void Rebind()
            {
                FigureBase.SuppressRenameInExpressions = false;
                foreach (var holder in expressions)
                {
                    holder.RebindExpressions();
                }
            }

            using (Transaction.Create(drawing.ActionManager, false))
            {
                drawing.ActionManager.RecordAction(new CallMethodAction(
                    () => FigureBase.SuppressRenameInExpressions = true,
                    Rebind));
                replacement.Visible = point.Visible;
                replacement.Locked = point.Locked;
                if (point.Style != null && point.Style != drawing.StyleManager.AssignDefaultStyle(point))
                {
                    replacement.Style = point.Style;
                }

                // the replacement takes over the point's label, if any: "Label new points" must
                // not give it one of its own, neither now nor on redo
                SuppressAutoLabelPoints(drawing, suppress: true);
                Add(drawing, replacement);
                SuppressAutoLabelPoints(drawing, suppress: false);

                // the label goes first: one still on the point when its other dependents move
                // is listed with the replacement as well (SubstituteWith), a second time, and
                // the entry left over after undo fails the consistency check
                var label = point.Label;
                if (label != null)
                {
                    ReplaceDependency(label, point, replacement);
                    var handOver = new CallMethodAction(
                        () =>
                        {
                            point.Label = null;
                            replacement.Label = label;
                        },
                        () =>
                        {
                            replacement.Label = null;
                            point.Label = label;
                        });
                    drawing.ActionManager.RecordAction(handOver);
                }

                ReplaceWithExisting(point, replacement);

                // in the point's place in the list (the Figure List, the file), not at the end
                MoveBefore(drawing, replacement, point);
                Remove(point);
                SetProperty(drawing.ActionManager, new PropertyValue("Name", replacement), point.Name);
                drawing.ActionManager.RecordAction(new CallMethodAction(
                    Rebind,
                    () => FigureBase.SuppressRenameInExpressions = true));
            }
        }

        /// <summary>
        /// Moves <paramref name="figure"/> in the drawing's list to just before
        /// <paramref name="before"/>, together with what it is built on that comes after that
        /// place (a Number made for it), so that the list stays in dependency order. Nothing
        /// is taken off the canvas: a move is not a removal and an insertion to the list.
        /// </summary>
        public static void MoveBefore(Drawing drawing, IFigure figure, IFigure before)
        {
            var figures = drawing.Figures;
            int target = figures.IndexOf(before);
            if (target < 0 || figures.IndexOf(figure) <= target)
            {
                return;
            }

            var moving = new HashSet<IFigure>();
            Collect(figure);
            var order = figures.Where(moving.Contains).ToArray();
            var oldIndices = new int[order.Length];
            drawing.ActionManager.RecordAction(new CallMethodAction(
                () =>
                {
                    for (int i = 0; i < order.Length; i++)
                    {
                        oldIndices[i] = figures.IndexOf(order[i]);
                        figures.Move(oldIndices[i], target + i);
                    }
                },
                () =>
                {
                    for (int i = order.Length - 1; i >= 0; i--)
                    {
                        figures.Move(target + i, oldIndices[i]);
                    }
                }));

            void Collect(IFigure item)
            {
                if (!moving.Add(item))
                {
                    return;
                }

                foreach (var dependency in item.Dependencies)
                {
                    if (figures.IndexOf(dependency) > target)
                    {
                        Collect(dependency);
                    }
                }
            }
        }

        /// <summary>Recorded, so that redo and undo pass through the same state</summary>
        static void SuppressAutoLabelPoints(Drawing drawing, bool suppress)
        {
            drawing.ActionManager.RecordAction(new CallMethodAction(
                () => PointBase.SuppressAutoLabelPoints = suppress,
                () => PointBase.SuppressAutoLabelPoints = !suppress));
        }

        public static void Move(Drawing drawing, IEnumerable<IMovable> moving, Point offset, IEnumerable<IFigure> toRecalculate)
        {
            if (drawing.ActionManager == null)
            {
                moving.Move(offset);
                MoveAction.Recalculate(drawing, toRecalculate);
                return;
            }

            var action = new MoveAction(drawing, moving, offset, toRecalculate);
            drawing.ActionManager.RecordAction(action);
        }

        public static void SetProperty(ActionManager actionManager, IValueProvider valueProvider, object value)
        {
            SetPropertyAction action = new SetPropertyAction(valueProvider, value);
            if (actionManager == null)
            {
                action.Execute();
            }
            else
            {
                actionManager.RecordAction(action);
            }
        }

        public static void AddItem<T>(ActionManager actionManager, ICollection<T> list, T item)
        {
            AddItemAction<T> action = new AddItemAction<T>(list.Add, i => list.Remove(i), item);
            actionManager.RecordAction(action);
        }

        public static void RemoveItem<T>(ActionManager actionManager, IList<T> list, T item)
        {
            var action = new RemoveItemAction<T>(list, item);
            actionManager.RecordAction(action);
        }

#if !PLAYER

        public static void Paste(Drawing drawing, string xmlContent)
        {
            var action = new PasteAction(
                drawing,
                xmlContent);
            drawing.ActionManager.RecordAction(action);
        }

#endif

        public static void InsertDependency(IFigure figure, int index, IFigure dependency)
        {
            var action = new CallMethodAction(
                () =>
                {
                    figure.InsertDependencyCore(index, dependency);
                },
                () =>
                {
                    figure.RemoveDependencyCore(index, dependency);
                });
            figure.Drawing.ActionManager.RecordAction(action);
        }

        public static void RemoveDependency(IFigure figure, IFigure dependency)
        {
            var index = figure.Dependencies.IndexOf(dependency);
            var action = new CallMethodAction(
                () =>
                {
                    figure.RemoveDependencyCore(index, dependency);
                },
                () =>
                {
                    figure.InsertDependencyCore(index, dependency);
                });
            figure.Drawing.ActionManager.RecordAction(action);
        }
    }
}
