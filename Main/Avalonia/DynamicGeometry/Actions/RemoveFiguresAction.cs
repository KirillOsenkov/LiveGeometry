using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry
{
    public class RemoveFiguresAction : GeometryAction
    {
        public RemoveFiguresAction(Drawing drawing, IEnumerable<IFigure> figures)
            : base(drawing)
        {
            Figures = figures.ToArray();
        }

        public IEnumerable<IFigure> Figures { get; set; }

        protected override void ExecuteCore()
        {
            Drawing.Figures.Remove(Figures);
        }

        protected override void UnExecuteCore()
        {
            Drawing.Figures.AddRange(Figures);
        }
    }
}
