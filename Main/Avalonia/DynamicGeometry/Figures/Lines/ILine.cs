namespace DynamicGeometry
{
    public interface ILine : IFigure, ILinearFigure
    {
        PointPair Coordinates { get; }
    }
}