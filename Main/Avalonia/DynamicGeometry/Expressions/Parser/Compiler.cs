using System;

namespace DynamicGeometry
{
    public class Compiler : ICompilerService
    {
        public static Compiler Instance { get; } = new Compiler();

        public IExpressionTreeEvaluatorProvider ExpressionTreeEvaluatorProvider { get; set; } = new ExpressionTreeCompiler();

        public CompileResult CompileFunction(Drawing drawing, string functionText)
        {
            CompileResult result = new CompileResult();
            if (string.IsNullOrEmpty(functionText))
            {
                return result;
            }

            Node ast = Parse(functionText, result);
            if (!result.Errors.IsEmpty())
            {
                return result;
            }

            ExpressionTreeBuilder builder = new ExpressionTreeBuilder();
            builder.SetContext(drawing, f => true);
            var expressionTree = builder.CreateFunction(ast, result);
            if (expressionTree == null || !result.Errors.IsEmpty())
            {
                return result;
            }

            Func<double, double> function = ExpressionTreeEvaluatorProvider.InterpretFunction(expressionTree);
            result.Function = function;
            return result;
        }

        public CompileResult CompileExpression(
            Drawing drawing,
            string expressionText,
            Predicate<IFigure> isFigureAllowed)
        {
            CompileResult result = new CompileResult();
            if (expressionText.IsEmpty())
            {
                return result;
            }

            Node ast = Parse(expressionText, result);
            if (!result.Errors.IsEmpty())
            {
                return result;
            }

            ExpressionTreeBuilder builder = new ExpressionTreeBuilder();
            builder.SetContext(drawing, isFigureAllowed);
            try
            {
                var expressionTree = builder.CreateExpression(ast, result);
                if (expressionTree == null || !result.Errors.IsEmpty())
                {
                    return result;
                }

                Func<double> function = ExpressionTreeEvaluatorProvider.InterpretExpression(expressionTree);
                result.Expression = function;
            }
            catch (Exception ex)
            {
                // an expression the builder can't make sense of (a function of an expression
                // where it expects a point, say) is an error of the label, not of the drawing
                result.AddError(ex.InnerException?.Message ?? ex.Message);
            }

            return result;
        }

        private static Node Parse(string text, CompileResult result)
        {
            ParseResult ast = Parser.Parse(text);
            if (!ast.Errors.IsEmpty())
            {
                result.Errors.AddRange(ast.Errors);
            }
            return ast.Root;
        }
    }
}
