using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using Irony.Parsing;

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

            ParseTree ast = Parse(functionText, result);
            if (!result.Errors.IsEmpty())
            {
                return result;
            }

            ExpressionTreeBuilder builder = new ExpressionTreeBuilder();
            builder.SetContext(drawing, f => true);
            var expressionTree = builder.CreateFunction(ast.Root, result);
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

            ParseTree ast = Parse(expressionText, result);
            if (!result.Errors.IsEmpty())
            {
                return result;
            }

            ExpressionTreeBuilder builder = new ExpressionTreeBuilder();
            builder.SetContext(drawing, isFigureAllowed);
            var expressionTree = builder.CreateExpression(ast.Root, result);
            if (expressionTree == null || !result.Errors.IsEmpty())
            {
                return result;
            }
            Func<double> function = ExpressionTreeEvaluatorProvider.InterpretExpression(expressionTree);
            result.Expression = function;
            return result;
        }

        private static ParseTree Parse(string text, CompileResult result)
        {
            ParseTree ast = ParserInstance.Parse(text);
            var compileErrors = GetCompilerErrors(ast.ParserMessages);
            if (!compileErrors.IsEmpty())
            {
                result.Errors.AddRange(compileErrors);
            }
            return ast;
        }

        private static List<CompileError> GetCompilerErrors(ParserMessageList syntaxErrors)
        {
            if (syntaxErrors.IsEmpty())
            {
                return null;
            }

            return syntaxErrors.Select(e => new CompileError()
            {
                Text = TranslateMessage(e)
            }).ToList();
        }

        private static string TranslateMessage(ParserMessage e)
        {
            var result = e.Message;
            if (result == "Unexpected end of file.")
            {
                result = "Unfinished expression";
            }
            return result;
        }

        private static Parser ParserInstance = new Parser(ExpressionGrammar.Instance);
    }
}
