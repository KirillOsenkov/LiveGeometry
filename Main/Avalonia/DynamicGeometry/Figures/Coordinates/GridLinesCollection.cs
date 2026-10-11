using Avalonia;

namespace DynamicGeometry
{
    public abstract class GridLinesCollection : FigureBase
    {
        public GridLinesCollection()
        {
            Layer = ZOrder.Grid;
        }

        public override IFigure HitTest(Point point)
        {
            return null;
        }
    }
}
