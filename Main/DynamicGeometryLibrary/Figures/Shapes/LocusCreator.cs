using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;

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
            }

            Drawing.RaiseConstructionStepStarted();

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

            Drawing.Figures.CheckConsistency();
        }

        /// <summary>
        /// It is important to exclude TempResults from the search since
        /// we don't want the figure to depend on its own parts.
        /// </summary>
        protected override IFigure LookForExpectedDependencyUnderCursor(Point coordinates)
        {
            return Drawing.Figures.HitTest(coordinates, f =>
            {
                if (f == null || !f.Visible || !f.IsHitTestVisible)
                {
                    return false;
                }

                if (FoundDependencies.Count == 0 && f.Dependencies.Count == 0)
                {
                    return false;
                }
                else if (FoundDependencies.Count == 1 && !((f is PointOnFigure) && f.Dependents.Contains(FoundDependencies[0])))
                {
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
                return "Click a point that depends on some point on figure.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Line(0, 0.7, 0.7, 0)
                .Line(0.3, 1, 1, 0.3)
                .Point(0.35, 0.35)
                .Canvas;
        }
    }
}
