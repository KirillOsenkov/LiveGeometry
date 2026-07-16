using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry
{
    /// <summary>
    /// Replaces a figure in a drawing with another figure
    /// </summary>
    public class ReplaceFigureAction : GeometryAction
    {
        public ReplaceFigureAction(Drawing drawing, IFigure figure, IFigure replacement)
            : base(drawing)
        {
            Figure = figure;
            Replacement = replacement;
        }

        public IFigure Figure { get; set; }
        public IFigure Replacement { get; set; }
        public IFigure[] Dependents { get; set; }

        protected override void ExecuteCore()
        {
            // I modified the following line to exclude PointLabels. When labeled points are replaced, they manage deleting and adding of labels.
            // Without filtering out PointLabels here, UnExecuteCore will result in an eventual CheckConsistency error. - D.H.
            Dependents = Figure.Dependents.Where(f => !(f is PointLabel)).ToArray();
            Figure.SubstituteWith(Replacement);
            RecalculateDependents();
        }

        void RecalculateDependents()
        {
            Drawing.Recalculate();
        }

        protected override void UnExecuteCore()
        {
            foreach (var dependent in Dependents)
            {
                dependent.ReplaceDependency(Replacement, Figure);
            }
            //Figure.Dependents.AddRange(Dependents);   This appears to be redundant and leads to errors. D.H.
            RecalculateDependents();
        }
    }
}
