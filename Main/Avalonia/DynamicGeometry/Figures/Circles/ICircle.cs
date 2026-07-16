using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;

namespace DynamicGeometry
{
    public interface ICircle : IEllipse
    {
        double Radius { get; }
    }
}