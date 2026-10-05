using System.Collections.Generic;

namespace DynamicGeometry
{
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
            Drawing.Figures.Add(Figure);
        }

        protected override void UnExecuteCore()
        {
            Drawing.Figures.Remove(Figure);
            AxisLine.Remove(Drawing, addedAxes);
        }
    }
}
