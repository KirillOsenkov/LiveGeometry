using System.Collections.Generic;

namespace DynamicGeometry
{
    public class Segment : LineBase, ILengthProvider, ILine
    {
        [PropertyGridVisible]
        public double Length
        {
            get
            {
                return Coordinates.Length;
            }
            set
            {
                var length = Length;
                if (value != length && Dependencies != null && Dependencies.Count > 1)
                {
                    var pointToMove = Dependencies[1] as IMovable;
                    if (pointToMove != null && pointToMove.AllowMove())
                    {
                        var factor = value / length;
                        var newLoc = Math.GetDilationPoint(Point(1), Point(0), factor);
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
            }
        }

        public override double GetNearestParameterFromPoint(System.Windows.Point point)
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

        public override IFigure HitTest(System.Windows.Point point)
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