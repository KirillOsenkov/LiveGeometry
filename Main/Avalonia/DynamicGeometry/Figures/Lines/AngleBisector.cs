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
                    if (Interior && result > 180)
                    {
                        result = 360 - result;
                    }
                }
                return result;
            }
        }

        /// <summary>
        /// Halves the angle under 180° between the sides, whichever way round they are. Off,
        /// the bisector halves the angle counterclockwise from the first side to the second, so
        /// it swings outside a triangle whose vertices get dragged the other way round. On for
        /// new bisectors; files from before have it off, as they were.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Inside the angle")]
        public bool Interior
        {
            get
            {
                return interior;
            }
            set
            {
                interior = value;
                if (Drawing != null)
                {
                    this.RecalculateAllDependents();
                }
            }
        }

        bool interior = true;

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
            interior = element.ReadBool("Interior", false);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            if (IsLine)
            {
                writer.WriteAttributeBool("Line", true);
            }

            if (Interior)
            {
                writer.WriteAttributeBool("Interior", true);
            }
        }

        /// <summary>
        /// The bisector of the other angle the two sides make (the one over 180°): the same
        /// line, pointing the other way. Inside-the-angle mode has no other angle, so it goes.
        /// </summary>
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
            interior = false;
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
                var vertex = dependencies.Point(0);
                var side1 = dependencies.Point(Flipped ? 2 : 1);
                var side2 = dependencies.Point(Flipped ? 1 : 2);

                // the halfway direction counterclockwise from side 1 to side 2; inside the angle
                // means the other way round when that sweep is the long way
                var halfway = Math.GetAngleBisectorPoint(vertex, side1, side2);
                if (Interior && halfway.Exists() && Math.OAngle(side1, vertex, side2) > Math.PI)
                {
                    halfway = vertex.Minus(halfway.Minus(vertex));
                }

                coordinates.P1 = vertex;
                coordinates.P2 = halfway;
                Exists = coordinates.P2.Exists();
            }
            else
            {
                Exists = false;
            }
        }
    }
}