using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;

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
