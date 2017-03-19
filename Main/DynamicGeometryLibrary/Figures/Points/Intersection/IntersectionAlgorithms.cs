using System.Windows;

namespace DynamicGeometry
{
    public class IntersectionAlgorithms
    {
        public static bool CanIntersect(IFigure figure1, IFigure figure2)
        {
            var lineEllipse = (figure1 is ILine || figure1 is IEllipse)
                && (figure2 is ILine || figure2 is IEllipse)
                && !(figure1 is IEllipse && figure2 is IEllipse);
            var lineCircle = (figure1 is ILine || figure1 is ICircle)
                && (figure2 is ILine || figure2 is ICircle);
            return lineEllipse || lineCircle;
        }

        #region Line and Line

        public static Point IntersectLineAndLine(IFigure line1, IFigure line2)
        {
            return Math.GetIntersectionOfLines(
                ((ILine)line1).Coordinates,
                ((ILine)line2).Coordinates);
        }

        #endregion

        #region Circle and Line

        /// <summary>
        /// The "Circle" methods are necessary since the serialization algorithm
        /// expects these to be here when reading .lgf files serialized with the
        /// old version.
        /// </summary>
        public static Point IntersectCircleAndLine1(IFigure ellipse, IFigure line)
        {
            return IntersectEllipseAndLine(ellipse, line).P1;
        }

        /// <summary>
        /// The "Circle" methods are necessary since the serialization algorithm
        /// expects these to be here when reading .lgf files serialized with the
        /// old version.
        /// </summary>
        public static Point IntersectCircleAndLine2(IFigure ellipse, IFigure line)
        {
            return IntersectEllipseAndLine(ellipse, line).P2;
        }

        /// <summary>
        /// The "Circle" methods are necessary since the serialization algorithm
        /// expects these to be here when reading .lgf files serialized with the
        /// old version.
        /// </summary>
        public static Point IntersectLineAndCircle1(IFigure line, IFigure ellipse)
        {
            return IntersectEllipseAndLine1(ellipse, line);
        }

        /// <summary>
        /// The "Circle" methods are necessary since the serialization algorithm
        /// expects these to be here when reading .lgf files serialized with the
        /// old version.
        /// </summary>
        public static Point IntersectLineAndCircle2(IFigure line, IFigure ellipse)
        {
            return IntersectEllipseAndLine2(ellipse, line);
        }

        #endregion

        #region Ellipse and Line

        public static PointPair IntersectEllipseAndLine(IFigure ellipse1, IFigure line1)
        {
            IEllipse ellipse = (IEllipse)ellipse1;
            ILine line = (ILine)line1;
            return Math.GetIntersectionOfEllipseAndLine(
                ellipse.Center,
                ellipse.SemiMajor,
                ellipse.SemiMinor,
                ellipse.Inclination,
                line.Coordinates);
        }

        public static Point IntersectEllipseAndLine1(IFigure ellipse, IFigure line)
        {
            return IntersectEllipseAndLine(ellipse, line).P1;
        }

        public static Point IntersectEllipseAndLine2(IFigure ellipse, IFigure line)
        {
            return IntersectEllipseAndLine(ellipse, line).P2;
        }

        public static PointPair IntersectLineAndEllipse(IFigure line, IFigure ellipse)
        {
            return IntersectEllipseAndLine(ellipse, line);
        }

        public static Point IntersectLineAndEllipse1(IFigure line, IFigure ellipse)
        {
            return IntersectEllipseAndLine1(ellipse, line);
        }

        public static Point IntersectLineAndEllipse2(IFigure line, IFigure ellipse)
        {
            return IntersectEllipseAndLine2(ellipse, line);
        }

        #endregion

        #region Circle and Circle

        public static PointPair IntersectCircleAndCircle(IFigure circle1, IFigure circle2)
        {
            ICircle c1 = (ICircle)circle1;
            ICircle c2 = (ICircle)circle2;
            return Math.GetIntersectionOfCircles(c1.Center, c1.Radius, c2.Center, c2.Radius);
        }

        public static Point IntersectCircleAndCircle1(IFigure circle1, IFigure circle2)
        {
            return IntersectCircleAndCircle(circle1, circle2).P1;
        }

        public static Point IntersectCircleAndCircle2(IFigure circle1, IFigure circle2)
        {
            return IntersectCircleAndCircle(circle1, circle2).P2;
        }

        #endregion
    }
}
