using System.Collections.Generic;
using System.Linq;
using GuiLabs.Undo;

namespace LiveGeometry;

/// <summary>
/// A figure that can keep going when some of what it is built on is deleted (a show/hide
/// box, a polygon, a polyline, a Bezier path) instead of going with it (<see cref="RemoveFigureAction"/>)
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

    /// <summary>
    /// Takes away all of those that go in one deletion: an action per dependency, unless the
    /// figure has to see them together (a polygon draws the side that closes the gap they
    /// leave). Made once and run again on redo, so they bring back the same figures.
    /// </summary>
    IEnumerable<IAction> GetRemoveDependencyActions(IList<IFigure> dependencies)
    {
        return dependencies.Select(GetRemoveDependencyAction).ToList();
    }

    /// <summary>
    /// What else the figure stops being built on when those go (the helper points of an
    /// image's handles, with the anchor whose handles they are): an auxiliary one goes in the
    /// same deletion once nothing else is built on it (<see cref="RemoveFigureAction.FindOrphanedAuxiliaries"/>)
    /// </summary>
    IEnumerable<IFigure> GetReleasedDependencies(IList<IFigure> dependencies)
    {
        return Enumerable.Empty<IFigure>();
    }
}
