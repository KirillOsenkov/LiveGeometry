using System.Windows;

namespace DynamicGeometry
{
    public static class ILinearFigureExtensions
    {
        public static Point SnapPointToFigure(this ILinearFigure figure, Point point)
        {
            var parameter = figure.GetNearestParameterFromPoint(point);
            return figure.GetPointFromParameter(parameter);
        }

        public static bool IsPointWithinTolerance(this ILinearFigure figure, Point point)
        {
            Point pointOnFigure = figure.SnapPointToFigure(point);
            return Math.Abs(pointOnFigure.Distance(point)) < figure.Drawing.CoordinateSystem.CursorTolerance;
        }

        public static ILinearFigure GetFigureIfPointWithinTolerance(this ILinearFigure figure, Point point)
        {            
            if (IsPointWithinTolerance(figure, point))
            {
                return figure;
            }
            else
            {
                return null;
            }
        }
    }
}