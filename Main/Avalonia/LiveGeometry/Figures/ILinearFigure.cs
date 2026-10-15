using Avalonia;

namespace LiveGeometry;

public interface ILinearFigure : IFigure
{
    double GetNearestParameterFromPoint(Point point);
    Point GetPointFromParameter(double parameter);
    Tuple<double, double> GetParameterDomain();
}