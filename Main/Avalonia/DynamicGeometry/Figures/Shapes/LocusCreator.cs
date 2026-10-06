using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Misc)]
    [Order(2)]
    public class LocusCreator : FigureCreator
    {
        protected override void AddDependency(Point coordinates)
        {
            IFigure underMouse = null;

            if (GetExpectedDependencyType() != null)
            {
                underMouse = LookForExpectedDependencyUnderCursor(coordinates);
                if (underMouse != null && FoundDependencies.Contains(underMouse) && !CanReuseDependency)
                {
                    return;
                }

                // nothing here that the locus takes: no step (see FigureCreator.AddDependency)
                if (underMouse == null || !GetExpectedDependencyType().IsAssignableFrom(underMouse.GetType()))
                {
                    return;
                }
            }

            StartConstruction();

            if (GetExpectedDependencyType() != null)
            {
                AddFoundDependency(underMouse);
            }

            if (GetExpectedDependencyType() != null)
            {
                AdvertiseNextDependency();
            }
            else
            {
                AddFiguresAndRestart();
            }

            Drawing.Figures.CheckConsistencyInDebug();
        }

        /// <summary>A click takes a point that is there and makes none: the choice is among those</summary>
        protected override IReadOnlyList<object> FindClickOptions(MouseEventArgs e)
        {
            return FindExpectedDependencies(Coordinates(e, false, false, false)).ToList<object>();
        }

        /// <summary>
        /// It is important to exclude TempResults from the search since
        /// we don't want the figure to depend on its own parts.
        /// </summary>
        protected override IReadOnlyList<IFigure> FindExpectedDependencies(Point coordinates)
        {
            return Drawing.Figures.HitTestAll(coordinates, f =>
            {
                if (f == null || !f.Visible || !f.IsHitTestVisible)
                {
                    return false;
                }

                if (FoundDependencies.Count == 0 && f.Dependencies.Count == 0)
                {
                    return false;
                }
                else if (FoundDependencies.Count == 1)
                {
                    if (f is PointOnFigure && f.AllDependents().Contains(FoundDependencies[0]))
                    {
                        return true;
                    }

                    return false;
                }

                return true;
            });
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            yield return Factory.CreateLocus(Drawing, FoundDependencies);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.Create<IPoint, PointOnFigure>();
        }

        protected override bool CanCreateTempResults()
        {
            return false;
        }

        public override string Name
        {
            get { return "Locus"; }
        }

        public override string HintText
        {
            get
            {
                return "Click a point built on a point that slides along a figure, then the sliding point: the locus is the path of the first.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            // The midpoint of a segment from a point that goes round a circle to a free point
            // goes round a circle of half the radius, centered halfway to the free point
            const string line = nameof(AppTheme.Line);
            Point center = new Point(0.36, 0.62);
            const double radius = 0.3;
            Point free = new Point(0.9, 0.1);
            Point onCircle = center + new Point(System.Math.Cos(0.9), System.Math.Sin(0.9)) * radius;
            Point locusCenter = (center + free) / 2;
            Point midpoint = (onCircle + free) / 2;
            return IconBuilder.BuildIcon()
                .Circle(strokeThickness: 1, line, center.X, center.Y, radius)
                .Circle(IconBuilder.AccentThickness, nameof(AppTheme.LineAccent), locusCenter.X, locusCenter.Y, radius / 2)
                .Line(line, onCircle.X, onCircle.Y, free.X, free.Y)
                .Point(free.X, free.Y)
                .Point(onCircle.X, onCircle.Y, nameof(AppTheme.PointOnFigureFill))
                .Point(midpoint.X, midpoint.Y, nameof(AppTheme.PointOnFigureFill))
                .Canvas;
        }
    }
}
