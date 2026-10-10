using System;
using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry
{
    public class AngleBisector : Ray, IConditionalProperties, IHasSweep
    {
        PointPair coordinates;

        // built on the angle's points, not through them: not "ABC"
        protected override IReadOnlyList<string> NamesFromDependencies()
        {
            return null;
        }

        protected override string Kind
        {
            get
            {
                return "Angle bisector";
            }
        }

        public override string Noun
        {
            get
            {
                return IsLine ? "line" : "ray";
            }
        }

        /// <summary>"of angle HIJ"</summary>
        public override string Construction
        {
            get
            {
                if (Dependencies.Count == 1)
                {
                    return "of " + ConstructionText.AngleValue(Dependencies[0]);
                }

                return Dependencies.Count == 3 ? "of " + ConstructionText.Angle(Dependencies[0], Dependencies[1], Dependencies[2]) : null;
            }
        }

        public override PointPair Coordinates
        {
            get
            {
                return coordinates;
            }
        }

        /// <summary>The angle halved, in degrees</summary>
        [PropertyGridVisible]
        public override double Angle
        {
            get
            {
                var dependencies = GetDependencies();
                return dependencies != null ? Sweep.Measure(CounterclockwiseAngle(dependencies)).ToDegrees() : 0;
            }
        }

        /// <summary>The counterclockwise angle from the first side to the second, 0 to 2π: what the sweep chooses from</summary>
        static double CounterclockwiseAngle(IFigure[] dependencies)
        {
            return Math.OAngle(dependencies.Point(1), dependencies.Point(0), dependencies.Point(2));
        }

        AngleSweep sweep = DefaultSweep;

        /// <summary>A new bisector halves the angle under 180°, whichever way round its sides were clicked</summary>
        public const AngleSweep DefaultSweep = AngleSweep.Smaller;

        /// <summary>
        /// Which of the two angles between the sides is halved (<see cref="AngleSweep"/>). A
        /// bisector built on an angle measurement halves the angle that says, and the row is
        /// read-only then (<see cref="CanEdit"/>).
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public AngleSweep Sweep
        {
            get
            {
                return Dependencies.Count == 1 && Dependencies[0] is IHasSweep angle ? angle.Sweep : sweep;
            }
            set
            {
                sweep = value;
                if (Drawing != null)
                {
                    this.RecalculateAllDependents();
                }
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
            sweep = element.ReadSweep(DefaultSweep);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            if (IsLine)
            {
                writer.WriteAttributeBool("Line", true);
            }

            // (a bisector of a measurement takes the measurement's sweep: nothing of its own to say)
            if (sweep != DefaultSweep && Dependencies.Count == 3)
            {
                writer.WriteAttributeString("Sweep", sweep.ToString());
            }
        }

        /// <summary>
        /// A ray's verbs that are not a bisector's: it is built on an angle, not on two points
        /// to draw a line or a segment through, and the way it points is the angle's
        /// (<see cref="Sweep"/>). The whole line is <see cref="IsLine"/>. The sweep of a
        /// bisector built on a measurement is the measurement's.
        /// </summary>
        public bool CanEdit(string propertyName)
        {
            if (propertyName == nameof(Sweep))
            {
                return Dependencies.Count == 3;
            }

            return propertyName != nameof(ConvertToLine)
                && propertyName != nameof(ConvertToSegment)
                && propertyName != nameof(Reverse);
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
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
                var side1 = dependencies.Point(1);
                var side2 = dependencies.Point(2);

                // the halfway direction counterclockwise from side 1 to side 2, or the other
                // way round when the sweep chooses the clockwise region
                var halfway = Math.GetAngleBisectorPoint(vertex, side1, side2);
                if (halfway.Exists() && Sweep.IsClockwise(CounterclockwiseAngle(dependencies)))
                {
                    halfway = vertex.Minus(halfway.Minus(vertex));
                }

                coordinates.P1 = vertex;
                coordinates.P2 = halfway;

                // (see ReflectedPoint: no bisector of an angle whose points are not there)
                Exists = Dependencies.Exists() && dependencies.Exists() && coordinates.P2.Exists();
            }
            else
            {
                Exists = false;
            }
        }
    }
}