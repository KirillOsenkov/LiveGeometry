using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Media;

namespace DynamicGeometry
{
    public class IntersectionPoint : PointBase, IPoint
    {
        public IntersectionPoint()
        {
        }

        public IntersectionPoint(Point hintPoint, IList<IFigure> dependencies)
        {
            Dependencies.AddRange(dependencies);
            IFigure figure1 = Dependencies.ElementAt(0);
            IFigure figure2 = Dependencies.ElementAt(1);
            Algorithm = DoubleDispatchIntersectionAlgorithm(figure1, figure2, hintPoint);
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            var algorithm = element.ReadString("Algorithm");
            if (string.IsNullOrEmpty(algorithm))
            {
                throw new Exception("When reading the IntersectionPoint, the Algorithm attribute "
                    + "was not specified. This point will not be created. Full text:\n"
                    + element.ToString());
            }
            var method = typeof(IntersectionAlgorithms).GetMethod(algorithm);
            if (method == null)
            {
                throw new Exception(string.Format("When reading the IntersectionPoint, the Algorithm method "
                    + "'{0}' wasn't found.", algorithm));
            }
            var @delegate = Delegate.CreateDelegate(typeof(Func<IFigure, IFigure, Point>), method);
            Algorithm = @delegate as Func<IFigure, IFigure, Point>;
            this.RecalculateAndUpdateVisual();
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeString("Algorithm", Algorithm.Method.Name);
        }

        protected override Avalonia.Controls.Shapes.Shape CreateShape()
        {
            var result = Factory.CreateDependentPointShape();
            result.Fill = new SolidColorBrush(Colors.Cyan);
            return result;
        }

        Func<IFigure, IFigure, Point> Algorithm;

        /// <summary>The name of the intersection algorithm, as saved in the Algorithm attribute</summary>
        public string AlgorithmName => Algorithm?.Method.Name;

        /// <summary>
        /// Math.GetIntersectionOfCircleAndLine once changed which of the two intersections comes
        /// first when the line passes through the center of the circle ("New code - preserves
        /// order"). In a drawing from before that, this point is the other one of the two:
        /// squares built with a perpendicular and a circle end up on the wrong side of their
        /// segment. Takes the other algorithm if that is the case here.
        /// </summary>
        /// <returns>Whether anything changed</returns>
        public bool UpgradeLegacyCircleAndLineOrder()
        {
            var line = Dependencies.OfType<ILine>().FirstOrDefault();
            var ellipse = Dependencies.OfType<IEllipse>().FirstOrDefault();
            if (line == null || ellipse == null || Algorithm == null)
            {
                return false;
            }

            var name = Algorithm.Method.Name;
            bool isFirst = name.EndsWith("1");
            if (!isFirst && !name.EndsWith("2"))
            {
                return false;
            }

            // the same test, with the same rounding, as the branch that changed
            var center = ellipse.Center;
            var projection = Math.GetProjectionPoint(center, line.Coordinates);
            if (!center.Exists() || !projection.Exists() || center.Distance(projection).Round(4) != 0)
            {
                return false;
            }

            var other = typeof(IntersectionAlgorithms).GetMethod(name.Substring(0, name.Length - 1) + (isFirst ? "2" : "1"));
            if (other == null)
            {
                return false;
            }

            Algorithm = (Func<IFigure, IFigure, Point>)Delegate.CreateDelegate(typeof(Func<IFigure, IFigure, Point>), other);
            return true;
        }

        public override void Recalculate()
        {
            // first assume we exist
            Exists = true;

            // if any of our dependencies don't exist, return
            UpdateExistence();
            if (!Exists)
            {
                return;
            }

            var figure1 = Dependencies.ElementAt(0);
            var figure2 = Dependencies.ElementAt(1);
            if (Algorithm == null)
            {
                Exists = false;
                return;
            }

            Point p = Algorithm(figure1, figure2);
            if (!p.Exists() || figure1.HitTest(p) == null || figure2.HitTest(p) == null)
            {
                Exists = false;
                return;
            }

            Exists = true;
            Coordinates = p;
        }

        public static Func<IFigure, IFigure, Point> DoubleDispatchIntersectionAlgorithm(
            IFigure figure1,
            IFigure figure2,
            Point hintPoint)
        {
            if (figure1 is ILine)
            {
                if (figure2 is ILine)
                {
                    return IntersectionAlgorithms.IntersectLineAndLine;
                }
                else if (figure2 is IEllipse)
                {
                    return PickCloserIntersectionPoint(
                        IntersectionAlgorithms.IntersectLineAndEllipse1,
                        IntersectionAlgorithms.IntersectLineAndEllipse2,
                        figure1,
                        figure2,
                        hintPoint);
                }
            }
            else if (figure1 is IEllipse)
            {
                if (figure2 is ILine)
                {
                    return PickCloserIntersectionPoint(
                        IntersectionAlgorithms.IntersectEllipseAndLine1,
                        IntersectionAlgorithms.IntersectEllipseAndLine2,
                        figure1,
                        figure2,
                        hintPoint);
                }
                else if (figure1 is ICircle && figure2 is ICircle)
                {
                    // The intersection of two ellipses is not supported yet (if ever).  Only two circles.
                    return PickCloserIntersectionPoint(
                        IntersectionAlgorithms.IntersectCircleAndCircle1,
                        IntersectionAlgorithms.IntersectCircleAndCircle2,
                        figure1,
                        figure2,
                        hintPoint);
                }
            }
            return null;
        }

        public static Func<IFigure, IFigure, Point> PickCloserIntersectionPoint(
            Func<IFigure, IFigure, Point> algorithm1,
            Func<IFigure, IFigure, Point> algorithm2,
            IFigure figure1,
            IFigure figure2,
            Point hintPoint)
        {
            Point p1 = algorithm1(figure1, figure2);
            Point p2 = algorithm2(figure1, figure2);

            if (!p1.Exists())
            {
                if (p2.Exists())
                {
                    return algorithm2;
                }
                else
                {
                    return algorithm1;
                }
            }
            else
            {
                if (!p2.Exists())
                {
                    return algorithm1;
                }
                else
                {
                    var d1 = p1.Distance(hintPoint);
                    var d2 = p2.Distance(hintPoint);
                    if (d1 < d2)
                    {
                        return algorithm1;
                    }
                    else
                    {
                        return algorithm2;
                    }
                }
            }
        }
    }
}

