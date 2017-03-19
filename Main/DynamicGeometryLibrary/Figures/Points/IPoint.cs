using System.Windows;

namespace DynamicGeometry
{
    public partial interface IPoint : IFigure
    {
        Point Coordinates { get; }
    }
}