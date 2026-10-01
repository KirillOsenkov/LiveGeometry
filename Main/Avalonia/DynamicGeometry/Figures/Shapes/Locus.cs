using System.Collections.Generic;
using Avalonia;
using Avalonia.Media;

namespace DynamicGeometry
{
    public class Locus : Curve, ILinearFigure
    {
        private List<IFigure> figuresToRecalculate;

        int StepCount
        {
            get
            {
                return 60;
            }
        }

        public Locus()
        {
            for (int i = 0; i < StepCount; i++)
            {
                pathSegments.Add(new LineSegment());
            }
        }

        /// <summary>
        /// What lies between the point that slides and the point traced, worked out anew
        /// each time: kept from the first time, it went on naming a figure that had since
        /// been replaced by another. No chain - the second point is not on a figure, or the
        /// first is not built on it - is a locus of nothing: no curve, rather than an
        /// exception from every move (PointSnapping keeps both points what they are; a
        /// file may still say anything).
        /// </summary>
        public override void Recalculate()
        {
            var point = mDependencies.Count > 0 ? mDependencies[0] as IPoint : null;
            var pointOnFigure = mDependencies.Count > 1 ? mDependencies[1] as PointOnFigure : null;
            figuresToRecalculate = point != null && pointOnFigure != null && point != pointOnFigure
                ? GetFiguresToRecalculate(pointOnFigure, point)
                : null;
        }

        protected override void OnDependenciesChanged()
        {
            figuresToRecalculate = null;
        }

        // the samples of the curve as GetPoints took them last: where the sliding point
        // was (its parameter on its figure) and where that put the traced point
        readonly List<double> sampleParameters = new List<double>();
        readonly List<Point> samplePoints = new List<Point>();

        bool HasChain
        {
            get { return figuresToRecalculate != null && figuresToRecalculate.Count > 0; }
        }

        /// <summary>Where the traced point is when the sliding point is at the parameter; the caller puts the sliding point back</summary>
        Point Trace(PointOnFigure sliding, IPoint traced, double parameter)
        {
            sliding.Parameter = parameter;
            for (int i = 0; i < figuresToRecalculate.Count; i++)
            {
                figuresToRecalculate[i].Recalculate();
            }

            return traced.Coordinates;
        }

        public override void GetPoints(List<Point> result)
        {
            sampleParameters.Clear();
            samplePoints.Clear();
            if (!HasChain)
            {
                return;
            }

            var steps = StepCount;
            result.Capacity = steps + 1;

            var point = mDependencies[0] as IPoint;
            var pointOnFigure = mDependencies[1] as PointOnFigure;
            var domain = pointOnFigure.LinearFigure.GetParameterDomain();
            var oldParameter = pointOnFigure.Parameter;

            // by the count, not by adding the step up: rounding decided whether the last
            // sample but one was taken, and the curve had 60 or 61 points from one move to
            // the next
            for (int i = 0; i <= steps; i++)
            {
                double lambda = i == steps ? domain.Item2 : domain.Item1 + (domain.Item2 - domain.Item1) * i / steps;
                var sample = Trace(pointOnFigure, point, lambda);
                result.Add(sample);
                sampleParameters.Add(lambda);
                samplePoints.Add(sample);
            }

            Trace(pointOnFigure, point, oldParameter);
        }

        #region A point on the locus

        // A point on the locus is where the traced point is for some place of the sliding
        // point, and that place - the sliding point's parameter on its own figure - is
        // what the point keeps. (It kept a fraction of the length of the drawn curve. But
        // how much of the locus is drawn depends on the view when the sliding point is on
        // a line: the visible part of the line, twice over. The point then slid along the
        // curve with every pan and zoom, and a saved drawing opened with it elsewhere.)

        public override double GetNearestParameterFromPoint(Point point)
        {
            if (samplePoints.Count == 0)
            {
                return 0;
            }

            // the nearest place on the curve as drawn, and the parameter in proportion
            // between the two samples it lies between
            double nearest = double.MaxValue;
            double parameter = sampleParameters[0];
            for (int i = 0; i < samplePoints.Count - 1; i++)
            {
                var from = samplePoints[i];
                var to = samplePoints[i + 1];
                double dx = to.X - from.X;
                double dy = to.Y - from.Y;
                double lengthSquared = dx * dx + dy * dy;
                double ratio = lengthSquared > 0
                    ? ((point.X - from.X) * dx + (point.Y - from.Y) * dy) / lengthSquared
                    : 0;
                ratio = System.Math.Max(0, System.Math.Min(1, ratio));
                double x = from.X + dx * ratio - point.X;
                double y = from.Y + dy * ratio - point.Y;
                double distance = x * x + y * y;

                // (a sample where the traced point was nowhere compares as false)
                if (distance < nearest)
                {
                    nearest = distance;
                    parameter = sampleParameters[i] + (sampleParameters[i + 1] - sampleParameters[i]) * ratio;
                }
            }

            return parameter;
        }

        /// <summary>
        /// Worked out, not read off the drawn curve: the sliding point is put at the
        /// parameter, what lies between it and the traced point is recalculated, and
        /// everything is put back - as for every sample of the curve.
        /// </summary>
        public override Point GetPointFromParameter(double parameter)
        {
            var point = mDependencies.Count > 0 ? mDependencies[0] as IPoint : null;
            var pointOnFigure = mDependencies.Count > 1 ? mDependencies[1] as PointOnFigure : null;
            if (point == null || pointOnFigure == null || !HasChain || !parameter.IsValidValue())
            {
                return Math.InfinitePoint;
            }

            var oldParameter = pointOnFigure.Parameter;
            var result = Trace(pointOnFigure, point, parameter);
            if (!point.Exists)
            {
                result = Math.InfinitePoint;
            }

            Trace(pointOnFigure, point, oldParameter);
            return result;
        }

        /// <summary>The parameters of the sliding point: those of the figure it slides on</summary>
        public override Tuple<double, double> GetParameterDomain()
        {
            var pointOnFigure = mDependencies.Count > 1 ? mDependencies[1] as PointOnFigure : null;
            return pointOnFigure != null ? pointOnFigure.LinearFigure.GetParameterDomain() : base.GetParameterDomain();
        }

        #endregion

        List<IFigure> GetFiguresToRecalculate(PointOnFigure pointOnFigure, IPoint dependentPoint)
        {
            return DependencyAlgorithms.FindImpactedDependencyChain(pointOnFigure, dependentPoint);
        }
    }
}
