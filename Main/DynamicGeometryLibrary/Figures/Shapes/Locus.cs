using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;

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

        public override void Recalculate()
        {
            if (figuresToRecalculate == null)
            {
                var point = mDependencies[0] as IPoint;
                var pointOnFigure = mDependencies[1] as PointOnFigure;
                figuresToRecalculate = GetFiguresToRecalculate(pointOnFigure, point);
            }
        }

        protected override void OnDependenciesChanged()
        {
            figuresToRecalculate = null;
        }

        public override void GetPoints(List<Point> result)
        {
            if (figuresToRecalculate == null)
            {
                return;
            }

            result.Capacity = StepCount + 1;

            var point = mDependencies[0] as IPoint;
            var pointOnFigure = mDependencies[1] as PointOnFigure;
            var figure = pointOnFigure.LinearFigure;
            var domain = figure.GetParameterDomain();
            var oldParameter = pointOnFigure.Parameter;
            var steps = StepCount;
            if (steps == 0)
            {
                return;
            }

            var step = (domain.Item2 - domain.Item1) / steps;
            var end = domain.Item2 - step;
            for (double lambda = domain.Item1; lambda < end; lambda += step)
            {
                pointOnFigure.Parameter = lambda;
                for (int i = 0; i < figuresToRecalculate.Count; i++)
                {
                    figuresToRecalculate[i].Recalculate();
                }
                result.Add(point.Coordinates);
            }

            pointOnFigure.Parameter = domain.Item2;
            for (int i = 0; i < figuresToRecalculate.Count; i++)
            {
                figuresToRecalculate[i].Recalculate();
            }
            result.Add(point.Coordinates);

            pointOnFigure.Parameter = oldParameter;
            for (int i = 0; i < figuresToRecalculate.Count; i++)
            {
                figuresToRecalculate[i].Recalculate();
            }
        }

        List<IFigure> GetFiguresToRecalculate(PointOnFigure pointOnFigure, IPoint dependentPoint)
        {
            return DependencyAlgorithms.FindImpactedDependencyChain(pointOnFigure, dependentPoint);
        }
    }
}
