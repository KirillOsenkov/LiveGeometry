using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;

namespace DynamicGeometry
{
    public interface ILinearFigure : IFigure
    {
        double GetNearestParameterFromPoint(Point point);
        Point GetPointFromParameter(double parameter);
        Tuple<double, double> GetParameterDomain();
    }
}