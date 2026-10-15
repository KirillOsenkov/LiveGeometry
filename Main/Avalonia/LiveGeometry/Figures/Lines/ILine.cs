namespace LiveGeometry;

public interface ILine : IFigure, ILinearFigure
{
    PointPair Coordinates { get; }
}