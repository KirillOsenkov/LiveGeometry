using System.Windows;

namespace DynamicGeometry
{
    public interface ICircle : IEllipse
    {
        double Radius { get; }
    }
}