using System;

namespace DynamicGeometry
{
    public interface ICompilerService
    {
        CompileResult CompileFunction(Drawing drawing, string functionText);
        CompileResult CompileExpression(
            Drawing drawing,
            string expressionText,
            Predicate<IFigure> isFigureAllowed);
    }
}
