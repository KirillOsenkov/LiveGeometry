using System.Windows;

namespace DynamicGeometry
{
    public interface IEllipse : IFigure, ILinearFigure
    {
        double SemiMajor { get; }
        double SemiMinor { get; }

        /// <summary>
        /// Angle of inclination from horizontal.
        /// </summary>
        double Inclination { get; }
    }
}