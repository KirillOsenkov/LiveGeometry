using System;
using System.Linq;

namespace DynamicGeometry
{
    public class AngleBisector : Ray
    {
        PointPair coordinates;

        public override PointPair Coordinates
        {
            get
            {
                return coordinates;
            }
        }

        [PropertyGridVisible]
        public override double Angle
        {
            get
            {
                var dependencies = GetDependencies();
                double result = 0;
                if (dependencies != null)
                {
                   
                    result = (Flipped) ? 
                        Math.OAngle(
                        dependencies.Point(2),
                        dependencies.Point(0),
                        dependencies.Point(1)).ToDegrees() :
                        Math.OAngle(
                        dependencies.Point(1),
                        dependencies.Point(0),
                        dependencies.Point(2)).ToDegrees();
                }
                return result;
            }
        }

        /// <summary>
        /// Extends the bisector in both directions (which is what DG's bisector was): the
        /// figure is then a line for hit testing, intersections and points on it.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Whole line")]
        public bool IsLine
        {
            get
            {
                return isLine;
            }
            set
            {
                isLine = value;
                if (Drawing != null)
                {
                    this.RecalculateAllDependents();
                }
            }
        }

        bool isLine;

        public override PointPair OnScreenCoordinates
        {
            get
            {
                if (!IsLine)
                {
                    return base.OnScreenCoordinates;
                }

                return Math.GetLineFromSegment(Coordinates, CanvasLogicalBorders);
            }
        }

        public override IFigure HitTest(Avalonia.Point point)
        {
            if (!IsLine)
            {
                return base.HitTest(point);
            }

            var epsilon = ToLogical(this.Shape.StrokeThickness) / 2 + CursorTolerance;
            return Math.IsPointOnLine(Coordinates, point, epsilon) ? this : null;
        }

        public override double GetNearestParameterFromPoint(Avalonia.Point point)
        {
            return IsLine ? Math.GetProjection(point, Coordinates).Ratio : base.GetNearestParameterFromPoint(point);
        }

        public override Tuple<double, double> GetParameterDomain()
        {
            if (!IsLine)
            {
                return base.GetParameterDomain();
            }

            var coordinates = OnScreenCoordinates;
            return new Tuple<double, double>(
                GetNearestParameterFromPoint(coordinates.P1) * 2,
                GetNearestParameterFromPoint(coordinates.P2) * 2);
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            isLine = element.ReadBool("Line", false);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            if (IsLine)
            {
                writer.WriteAttributeBool("Line", true);
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to opposite angle")]
        public void ConvertToOpposite()
        {
            IFigure[] dependencies = Dependencies as IFigure[];
            if (dependencies != null)
            {
                var t = dependencies[1];
                dependencies[1] = dependencies[2];
                dependencies[2] = t;
            }
            this.RecalculateAllDependents();
            Drawing.RaiseSelectionChanged(this);
        }

        IFigure[] GetDependencies()
        {
            var dependencies = Dependencies.ToArray();
            if (dependencies.Length == 1)
            {
                AngleMeasurement angle = dependencies[0] as AngleMeasurement;
                if (angle == null)
                {
                    return null;
                }
                dependencies = angle.Dependencies.ToArray();
            }
            if (dependencies.Length != 3)
            {
                dependencies = null;
            }
            return dependencies;
        }

        public override void Recalculate()
        {
            var dependencies = GetDependencies();
            if (dependencies != null)
            {
                coordinates.P1 = dependencies.Point(0);
                coordinates.P2 = (Flipped) ?
                    Math.GetAngleBisectorPoint(dependencies.Point(0), dependencies.Point(2), dependencies.Point(1)) :
                    Math.GetAngleBisectorPoint(dependencies.Point(0), dependencies.Point(1), dependencies.Point(2));
                Exists = coordinates.P2.Exists();
            }
            else
            {
                Exists = false;
            }
        }
    }
}