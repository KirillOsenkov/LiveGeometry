using System.Windows;

namespace DynamicGeometry
{
    public abstract class GridLinesCollection : FigureBase
    {
        public GridLinesCollection()
        {
            ZIndex = (int)ZOrder.Grid;
        }

        public override IFigure HitTest(Point point)
        {
            return null;
        }
    }
}
