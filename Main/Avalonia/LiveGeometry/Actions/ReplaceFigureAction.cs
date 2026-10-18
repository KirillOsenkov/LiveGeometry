using System.Linq;

namespace LiveGeometry;

/// <summary>
/// Replaces a figure in a drawing with another figure
/// </summary>
public class ReplaceFigureAction : GeometryAction
{
    public ReplaceFigureAction(Drawing drawing, IFigure figure, IFigure replacement)
        : base(drawing)
    {
        Figure = figure;
        Replacement = replacement;
    }

    public IFigure Figure { get; set; }
    public IFigure Replacement { get; set; }
    public IFigure[] Dependents { get; set; }

    protected override void ExecuteCore()
    {
        // a point's name goes over like any other dependent, as a line's does: the label is the
        // replacement's from then on (Actions.HandOverNameLabel)
        Dependents = Figure.Dependents.ToArray();
        Figure.SubstituteWith(Replacement);
        RecalculateDependents();
    }

    void RecalculateDependents()
    {
        Drawing.Recalculate();
    }

    protected override void UnExecuteCore()
    {
        foreach (var dependent in Dependents)
        {
            // a part of a composite has gone back with its composite already
            if (dependent.Dependencies.Contains(Replacement))
            {
                dependent.ReplaceDependency(Replacement, Figure);
            }
        }
        //Figure.Dependents.AddRange(Dependents);   This appears to be redundant and leads to errors. D.H.
        RecalculateDependents();
    }
}
