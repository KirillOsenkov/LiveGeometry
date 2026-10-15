using System;

namespace LiveGeometry;

public interface ICompilerService
{
    CompileResult CompileFunction(Drawing drawing, string functionText, Predicate<IFigure> isFigureAllowed = null);
    CompileResult CompileExpression(
        Drawing drawing,
        string expressionText,
        Predicate<IFigure> isFigureAllowed);
}
