using System.Collections.Generic;

namespace DynamicGeometry
{
    public class Segment : LineBase, ILengthProvider, ILine, IConditionalProperties
    {
        /// <summary>
        /// Setting it stretches the segment once, to that length, by moving an end away from
        /// the other: the second end if it can take it, else the first. Not a constraint -
        /// dragging changes the length again. Read-only when neither end can take it.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Length
        {
            get
            {
                return Coordinates.Length;
            }
            set
            {
                var length = Length;
                var end = EndToStretch();
                if (value == length || end < 0)
                {
                    return;
                }

                var pointToMove = (IMovable)Dependencies[end];
                var factor = value / length;
                var newLoc = Math.GetDilationPoint(Point(end), Point(1 - end), factor);
                pointToMove.MoveTo(newLoc);
                (pointToMove as IFigure).RecalculateAndUpdateVisual();
                List<IFigure> dependents = DependencyAlgorithms.FindDescendants(f => f.Dependents, (pointToMove as IFigure).AsEnumerable());
                dependents.Reverse();
                foreach (var dependent in dependents)
                {
                    dependent.RecalculateAndUpdateVisual();
                }
            }
        }

        /// <summary>
        /// The index of the end a new length moves, -1 when neither can go where the length
        /// says: a free point can, and so can a point that slides along this very segment's
        /// line (a translated point from the other end with a free magnitude). A point on
        /// some other figure would only get near, and an end the other end is built on
        /// (a fixed-length segment: the far end follows) would take the whole segment along.
        /// </summary>
        int EndToStretch()
        {
            if (Dependencies == null || Dependencies.Count < 2)
            {
                return -1;
            }

            for (int end = 1; end >= 0; end--)
            {
                if (CanTakeLength(Dependencies[end], Dependencies[1 - end]))
                {
                    return end;
                }
            }

            return -1;
        }

        static bool CanTakeLength(IFigure point, IFigure otherEnd)
        {
            if (point.Locked || otherEnd.DependsOn(point))
            {
                return false;
            }

            if (point is FreePoint)
            {
                return true;
            }

            return point is TranslatedPoint translated
                && translated.IsMagnitudeFree
                && !translated.IsDirectionFree
                && translated.Source == otherEnd;
        }

        public bool CanEdit(string propertyName)
        {
            return propertyName != "Length" || EndToStretch() >= 0;
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

        public override double GetNearestParameterFromPoint(Avalonia.Point point)
        {
            var parameter = base.GetNearestParameterFromPoint(point);
            if (parameter < 0)
            {
                parameter = 0;
            }
            else if (parameter > 1)
            {
                parameter = 1;
            }
            return parameter;
        }

        public override IFigure HitTest(Avalonia.Point point)
        {
            var epsilon = ToLogical(this.Shape.StrokeThickness) / 2 + CursorTolerance;
            if (Math.IsPointOnSegment(Coordinates, point, epsilon))
            {
                return this;
            }
            return null;
        }

        public override Tuple<double, double> GetParameterDomain()
        {
            return Tuple.Create(0.0, 1.0);
        }

        public override string ToString()
        {
            // I think it is confusing to the user when the title of the property grid for a segment is different than the name.
            // For example, the user might use the ToString() instead of the Name when referring to the segment in an expression. - D.H.
            return base.ToString();
            //return "Segment " + Dependencies[0].ToString() + Dependencies[1].ToString();
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to line")]
        public void ConvertToLine()
        {
            LineTwoPoints.Convert(this, Factory.CreateLineTwoPoints(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to ray")]
        public void ConvertToRay()
        {
            LineTwoPoints.Convert(this, Factory.CreateRay(this.Drawing, this.Dependencies));
        }

#endif
    }
}