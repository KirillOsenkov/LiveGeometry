using System.Linq;
using System.ComponentModel.Composition;

namespace DynamicGeometry
{
    [Export(typeof(IOperationProvider))]
    public class TearOffMutation : IOperationProvider
    {
        [PropertyGridName("Tear off")]
        public static void TearOffFigure(IFigure figure)
        {
            if (TearOff(figure as ParallelLine))
            {
                return;
            }

            TearOffGeneralCase(figure);
        }

        private static void TearOffGeneralCase(IFigure figure)
        {
            Drawing drawing = figure.Drawing;

            using (drawing.ActionManager.CreateTransaction())
            {
                foreach (var dependency in figure.Dependencies.ToArray())
                {
                    var dependencyPoint = dependency as IPoint;
                    if (dependencyPoint == null)
                    {
                        // for now, can't tear off a figure 
                        // that has a non-point dependency (such as a ParallelLine)
                        return;
                    }

                    if (dependencyPoint is FreePoint)
                    {
                        // no need to tear-off from an already free point
                        // since no one else uses this point
                        if (dependencyPoint.Dependents.Count == 1) continue;
                        if (dependencyPoint.Dependents.Count == 2 && dependencyPoint.Dependents.OfType<LabelBase>().Count() > 0) continue;
                    }

                    FreePoint newDependencyPoint = Factory.CreateFreePoint(drawing, dependencyPoint.Coordinates);
                    Actions.Add(figure.Drawing, newDependencyPoint);
                    Actions.ReplaceDependency(figure, dependencyPoint, newDependencyPoint);
                }
            }

            drawing.RaiseDisplayProperties(figure);
        }

        private static bool TearOff(ParallelLine parallelLine)
        {
            if (parallelLine == null)
            {
                return false;
            }

            Drawing drawing = parallelLine.Drawing;
            PointPair coordinates = parallelLine.Coordinates;
            FreePoint point1 = Factory.CreateFreePoint(drawing, coordinates.P1);
            FreePoint point2 = Factory.CreateFreePoint(drawing, coordinates.P2);
            LineTwoPoints line = Factory.CreateLineTwoPoints(drawing, new[] { point1, point2 });

            using (drawing.ActionManager.CreateTransaction())
            {
                Actions.Add(drawing, point1);
                Actions.Add(drawing, point2);
                Actions.Add(drawing, line);
                Actions.ReplaceWithExisting(parallelLine, line);
                Actions.Remove(parallelLine);
            }

            drawing.RaiseDisplayProperties(line);

            return true;
        }

        private static bool CanTearOff(IFigure figure)
        {
            if (figure is ParallelLine)
            {
                return true;
            }

            foreach (var dependency in figure.Dependencies)
            {
                var dependencyPoint = dependency as IPoint;
                if (dependencyPoint == null)
                {
                    // for now, can't tear off a figure 
                    // that has a non-point dependency (such as a ParallelLine)
                    return false;
                }

                if (dependencyPoint is FreePoint)
                {
                    // no need to tear-off from an already free point
                    // since no one else uses this point
                    if (dependencyPoint.Dependents.Count == 1) continue;
                    if (dependencyPoint.Dependents.Count == 2 && dependencyPoint.Dependents.OfType<LabelBase>().Count() > 0) continue;
                }

                return true;
            }

            return false;
        }

        public IOperationDescription ProvideOperation(object instance)
        {
            var figure = instance as IFigure;
            if (figure == null || !CanTearOff(figure))
            {
                return null;
            }

            return MethodDescription.Get<TearOffMutation>("TearOffFigure");
        }
    }
}
