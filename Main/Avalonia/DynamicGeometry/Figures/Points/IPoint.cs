using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;

namespace DynamicGeometry
{
    public partial interface IPoint : IFigure
    {
        Point Coordinates { get; }
    }
}