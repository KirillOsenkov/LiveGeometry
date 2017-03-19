using GuiLabs.Undo;
using System.Collections.Generic;
using System.Windows;

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
                if (newFigure is PointBase && existingFigure is PointBase)
                {
                    Actions.SetProperty(drawing.ActionManager, new PropertyValue("Name", newFigure), existingFigure.Name);
                }
            }
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

        public static void RemoveMany(Drawing drawing, IEnumerable<IFigure> figures)
        {
            // TODO: switch to using RemoveFigure multiple times
            var action = new RemoveFiguresAction(drawing, figures);
            drawing.ActionManager.RecordAction(action);
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
