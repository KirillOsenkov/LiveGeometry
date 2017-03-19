using System;
using System.ComponentModel.Composition;

namespace DynamicGeometry
{
    [Export(typeof(ICompilerService))]
    public class Compiler : ICompilerService
    {
        [Import]
        public IExpressionTreeEvaluatorProvider ExpressionTreeEvaluatorProvider { get; set; }

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
            var expressionTree = builder.CreateExpression(ast, result);
            if (expressionTree == null || !result.Errors.IsEmpty())
            {
                return result;
            }
            Func<double> function = ExpressionTreeEvaluatorProvider.InterpretExpression(expressionTree);
            result.Expression = function;
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
