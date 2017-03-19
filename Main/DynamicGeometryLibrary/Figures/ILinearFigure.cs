using System.Windows;

namespace DynamicGeometry
{
    public interface ILinearFigure : IFigure
    {
        double GetNearestParameterFromPoint(Point point);
        Point GetPointFromParameter(double parameter);
        Tuple<double, double> GetParameterDomain();
    }
}