using System.Collections.Generic;

namespace LiveGeometry;

public class AddFigureAction : GeometryAction
{
    public AddFigureAction(Drawing drawing, IFigure figure)
        : base(drawing)
    {
        Figure = figure;
    }

    public IFigure Figure { get; set; }

    // the axis lines the figure brought into the drawing, built on one of them
    List<AxisLine> addedAxes;

    protected override void ExecuteCore()
    {
        addedAxes = AxisLine.AddMissing(Drawing, new[] { Figure });
        try
        {
            Drawing.Figures.Add(Figure);
        }
        catch
        {
            // A figure that throws as it comes in (styled, a circle read its radius off a
            // line a defined tool had taken for a segment) is in the list by then, but to
            // the history the action never ran: undo and Escape left it there, and it was
            // saved. It goes again before the exception goes on.
            Drawing.Figures.Remove(Figure);
            AxisLine.Remove(Drawing, addedAxes);
            throw;
        }
    }

    protected override void UnExecuteCore()
    {
        Drawing.Figures.Remove(Figure);
        AxisLine.Remove(Drawing, addedAxes);
    }
}
