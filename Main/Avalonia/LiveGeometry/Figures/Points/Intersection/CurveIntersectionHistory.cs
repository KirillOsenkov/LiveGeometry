using System;
using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>
/// Undo and redo put a crossing with a curve back on its own crossing. Such a point follows
/// the crossing nearest where it was last (<see cref="IntersectionPoint.Parameter"/>), and
/// undo jumps: a parabola dragged far to the right and back by Ctrl+Z took both of its
/// crossings with the x-axis onto the one nearest where they had got to. So the parameters
/// are kept for every state of the history, by the action that led to it, and a state
/// reached again by undo or redo gets them back (a curve's crossings only: every other
/// point is worked out from the figures alone).
/// </summary>
public class CurveIntersectionHistory
{
    readonly Drawing drawing;

    // the state before the first action
    static readonly object Start = new();

    readonly Dictionary<object, Dictionary<IntersectionPoint, double>> states = new();

    object current = Start;

    public CurveIntersectionHistory(Drawing drawing)
    {
        this.drawing = drawing;
        states[Start] = Capture();
        drawing.ActionManager.CollectionChanged += ActionManager_CollectionChanged;
    }

    Dictionary<IntersectionPoint, double> Capture()
    {
        var result = new Dictionary<IntersectionPoint, double>();
        foreach (var point in drawing.Figures.OfType<IntersectionPoint>())
        {
            if (point.IsOnCurve)
            {
                result[point] = point.Parameter;
            }
        }

        return result;
    }

    void Restore(Dictionary<IntersectionPoint, double> state)
    {
        foreach (var (point, parameter) in state)
        {
            if (point.Parameter.Equals(parameter) || !drawing.Figures.Contains(point))
            {
                continue;
            }

            point.Parameter = parameter;
            point.RecalculateAndUpdateVisual();
            point.RecalculateAllDependents();
        }
    }

    /// <summary>A transaction rolled back (a construction given up): the parameters as they were before it</summary>
    public void RestoreCurrent()
    {
        if (states.TryGetValue(current, out var state))
        {
            Restore(state);
        }
    }

    void ActionManager_CollectionChanged(object sender, EventArgs e)
    {
        var manager = drawing.ActionManager;
        if (!manager.CanUndo && !manager.CanRedo)
        {
            // a new history: a drawing loaded, or a new one
            states.Clear();
            current = Start;
            states[Start] = Capture();
            return;
        }

        object reached = (object)manager.EnumUndoableActions().LastOrDefault() ?? Start;
        if (reached != current && states.TryGetValue(reached, out var state))
        {
            // undone or redone
            Restore(state);
        }
        else
        {
            // a new action, or one merged into the last (a run of moves)
            states[reached] = Capture();
        }

        current = reached;
    }
}
