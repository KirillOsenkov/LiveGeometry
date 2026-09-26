using System.Collections.Generic;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public class Segment : LineBase, ILengthProvider, ILine, IConditionalProperties
    {
        /// <summary>
        /// Setting it on a segment with a fixed length (<see cref="FixLength"/>) changes that
        /// length. Otherwise it stretches the segment once, to that length, by moving an end
        /// away from the other: the second end if it can take it, else the first - not a
        /// constraint, dragging changes the length again. Read-only when neither applies.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridGroup("Length")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Length
        {
            get
            {
                return Coordinates.Length;
            }
            set
            {
                var fixedEnd = FixedEnd();
                if (fixedEnd != null)
                {
                    fixedEnd.Distance = value;
                    return;
                }

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
        /// line (a translated point from the other end with a free distance). A point on
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
                && translated.IsDistanceFree
                && !translated.IsDirectionFree
                && translated.Source == otherEnd;
        }

        /// <summary>
        /// The end that holds the segment's length: a translated point from the other end
        /// whose distance is a Number. Null for a segment without a fixed length.
        /// </summary>
        TranslatedPoint FixedEnd()
        {
            for (int end = 1; end >= 0; end--)
            {
                if (Dependencies.Count > 1
                    && Dependencies[end] is TranslatedPoint translated
                    && translated.Source == Dependencies[1 - end]
                    && translated.DistanceSource is Number)
                {
                    return translated;
                }
            }

            return null;
        }

        public bool CanEdit(string propertyName)
        {
            switch (propertyName)
            {
                case "Length":
                    return FixedEnd() != null || EndToStretch() >= 0;
                case "FixLength":
                    return FixedEnd() == null && EndToStretch() >= 0;
                case "FreeLength":
                    return FixedEnd() != null;
                default:
                    return true;
            }
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

#if !PLAYER && !TABULA

        /// <summary>
        /// The segment keeps its length from now on: the end that could take a new length
        /// (see <see cref="EndToStretch"/>) becomes a translated point from the other end, at
        /// the current length held by a Number, direction free - so it drags around the other
        /// end and the other end carries it along. An end already sliding along the segment's
        /// line just has its distance fixed. Nothing moves.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Fix length")]
        [PropertyGridGroup("Length")]
        [PropertyGridIcon(PropertyGridIcon.Lock)]
        public void FixLength()
        {
            int end = EndToStretch();
            if (end < 0)
            {
                return;
            }

            var point = Dependencies[end];
            if (point is TranslatedPoint sliding)
            {
                Drawing.ActionManager.SetProperty(sliding, "FreeDistance", false);
            }
            else
            {
                var pivot = (IPoint)Dependencies[1 - end];
                using (Transaction.Create(Drawing.ActionManager, false))
                {
                    var length = Number.CreateAuxiliary(Drawing, Length);
                    Actions.Add(Drawing, length);
                    var fixedEnd = Factory.CreateTranslatedPoint(Drawing, pivot, length, directionSource: null);
                    // the free direction: where the end is now
                    fixedEnd.MoveTo(((IPoint)point).Coordinates);
                    Actions.ReplacePoint((PointBase)point, fixedEnd);
                }
            }

            Drawing.RaiseDisplayProperties(this);
        }

        /// <summary>
        /// The opposite of <see cref="FixLength"/>: the end holding the length becomes a free
        /// point where it is (its Number goes with it), or, when its direction is fixed too,
        /// keeps sliding along the line with the distance free.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Free length")]
        [PropertyGridGroup("Length")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void FreeLength()
        {
            var fixedEnd = FixedEnd();
            if (fixedEnd == null)
            {
                return;
            }

            if (fixedEnd.IsDirectionFree)
            {
                var free = Factory.CreateFreePoint(Drawing, fixedEnd.Coordinates);
                Actions.ReplacePoint(fixedEnd, free);
            }
            else
            {
                Drawing.ActionManager.SetProperty(fixedEnd, "FreeDistance", true);
            }

            Drawing.RaiseDisplayProperties(this);
        }

#endif

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
        [PropertyGridIcon(PropertyGridIcon.Line)]
        public void ConvertToLine()
        {
            LineTwoPoints.Convert(this, Factory.CreateLineTwoPoints(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to ray")]
        [PropertyGridIcon(PropertyGridIcon.Ray)]
        public void ConvertToRay()
        {
            LineTwoPoints.Convert(this, Factory.CreateRay(this.Drawing, this.Dependencies));
        }

#endif
    }
}