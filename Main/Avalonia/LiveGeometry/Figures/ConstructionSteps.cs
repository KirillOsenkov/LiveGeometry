using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>
/// How a figure takes part in stepping through the construction (<see cref="StepPlayback"/>):
/// a step of its own, in one step with its neighbors in the figure list, or left out.
/// </summary>
public enum StepRole
{
    [PropertyGridName("Own step")]
    Regular,

    [PropertyGridName("Group with next")]
    GroupWithNext,

    [PropertyGridName("Group with previous")]
    GroupWithPrevious,

    [PropertyGridName("Not a step")]
    Exclude
}

/// <summary>One step of the construction: a run of figures of the list shown together, and what the step says</summary>
public class ConstructionStep
{
    public List<IFigure> Figures { get; } = new List<IFigure>();

    /// <summary>
    /// The descriptions the author typed on the figures of the step, if any, else the
    /// sentence each figure says of itself
    /// </summary>
    public string Text
    {
        get
        {
            var figures = Figures.OfType<FigureBase>().ToList();
            var typed = figures
                .Select(figure => figure.Description)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();
            var parts = typed.Count > 0 ? typed : figures.Select(figure => figure.DefaultDescription).ToList();
            return string.Join(" ", parts.Select(part => part.Trim()));
        }
    }
}

/// <summary>
/// The steps of a drawing, read off the figure list: every visible figure that can be a
/// step is one, in the order of the list, and two neighbors are one step when the earlier
/// says <see cref="StepRole.GroupWithNext"/> or the later <see cref="StepRole.GroupWithPrevious"/>.
/// A point's name goes with its point; a caption, a show/hide box and the grid are always
/// on the paper (<see cref="FigureBase.CanBeStep"/>).
/// </summary>
public static class ConstructionSteps
{
    public static bool Participates(IFigure figure)
    {
        return figure is FigureBase figureBase
            && figureBase.CanBeStep()
            && figureBase.StepRole != StepRole.Exclude
            && figure.Visible;
    }

    public static List<ConstructionStep> Partition(Drawing drawing)
    {
        var steps = new List<ConstructionStep>();
        ConstructionStep current = null;
        FigureBase previous = null;
        foreach (var figure in drawing.Figures)
        {
            if (!Participates(figure))
            {
                continue;
            }

            var figureBase = (FigureBase)figure;
            bool bonded = current != null
                && (previous.StepRole == StepRole.GroupWithNext || figureBase.StepRole == StepRole.GroupWithPrevious);
            if (!bonded)
            {
                current = new ConstructionStep();
                steps.Add(current);
            }

            current.Figures.Add(figure);
            previous = figureBase;
        }

        return steps;
    }
}
