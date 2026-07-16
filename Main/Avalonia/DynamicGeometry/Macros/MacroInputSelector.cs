using System.Linq;

namespace DynamicGeometry
{
    [Ignore]
    public class MacroInputSelector : FigureSelector
    {
        protected override bool CanSelectFigure(IFigure figure)
        {
            if (DependencyAlgorithms.FigureCompletelyDependsOnFigures(figure, GetSelection().Without(figure)))
            {
                return false;
            }
            return true;
        }

        protected override void TrySelectFigure(IFigure figure)
        {
            if (!CanSelectFigure(figure))
            {
                return;
            }
            SelectFigure(figure);
            foreach (var selected in GetSelection())
            {
                if (DependencyAlgorithms.FigureCompletelyDependsOnFigures(selected, GetSelection().Without(selected)))
                {
                    DeselectFigure(selected);
                }
            }
        }
    }
}
