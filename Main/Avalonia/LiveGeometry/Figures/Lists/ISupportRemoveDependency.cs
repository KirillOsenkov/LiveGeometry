using System.Collections.Generic;
using GuiLabs.Undo;

namespace LiveGeometry;

/// <summary>
/// A figure that can keep going when some of what it is built on is deleted (a show/hide
/// box, a polyline, a Bezier path) instead of going with it (<see cref="RemoveFigureAction"/>)
/// </summary>
public interface ISupportRemoveDependency
{
    /// <summary>
    /// Whether the figure keeps going without those of its dependencies that are among
    /// the removed ones: everything the deletion takes, decided as a whole, since figures
    /// it holds may go together (a point and a segment on it)
    /// </summary>
    bool CanRemoveDependencies(ISet<IFigure> removed);

    /// <summary>
    /// Takes one dependency away. Several actions for one figure are made before any of
    /// them runs, so each must find its dependency when it runs, not when it is made.
    /// </summary>
    IAction GetRemoveDependencyAction(IFigure dependency);
}
