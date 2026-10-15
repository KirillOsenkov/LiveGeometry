using Avalonia;

namespace LiveGeometry;

public partial interface IPoint : IFigure
{
    Point Coordinates { get; }
}