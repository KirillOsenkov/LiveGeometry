using System;
using System.Diagnostics;
using System.Linq.Expressions;

namespace DynamicGeometry
{
    public class Compiler : ICompilerService
    {
        public static Compiler Instance { get; } = new Compiler();

        /// <summary>
        /// How the bound tree of an expression becomes the delegate that is run at every
        /// recalculation (<see cref="ExpressionStrategy"/>): the tree evaluating itself by
        /// default; a System.Linq.Expressions tree interpreted or compiled otherwise. Taken
        /// as each expression compiles, so a change applies to what compiles from then on.
        /// </summary>
        public ExpressionStrategy Strategy { get; set; } = ExpressionStrategy.OwnTree;

        /// <summary>
        /// The time spent compiling (parsing, binding, making the delegate) since this was
        /// last set: how much of a drawing's load its expressions are (ExpressionBenchmark)
        /// </summary>
        public TimeSpan CompileTime { get; set; }

        static readonly ExpressionTreeCompiler lightCompiler = new ExpressionTreeCompiler(preferInterpretation: true);
        static readonly ExpressionTreeCompiler ilCompiler = new ExpressionTreeCompiler(preferInterpretation: false);
        static readonly ExpressionTreeInterpreter interpreter = new ExpressionTreeInterpreter();

        IExpressionTreeEvaluatorProvider Provider
        {
            get
            {
                switch (Strategy)
                {
                    case ExpressionStrategy.LightCompiler:
                        return lightCompiler;
                    case ExpressionStrategy.LinqInterpreter:
                        return interpreter;
                    case ExpressionStrategy.Compiled:
                        return ilCompiler;
                    default:
                        throw new InvalidOperationException("The bound tree evaluates itself under " + Strategy + ".");
                }
            }
        }

        public CompileResult CompileFunction(Drawing drawing, string functionText, Predicate<IFigure> isFigureAllowed = null)
        {
            long started = Stopwatch.GetTimestamp();
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
            builder.SetContext(drawing, isFigureAllowed);
            try
            {
                var bound = builder.CreateFunction(ast, result);
                if (bound == null || !result.Errors.IsEmpty())
                {
                    return result;
                }

                result.Function = MakeFunction(bound);
            }
            catch (Exception ex)
            {
                // a tree the strategy can't make a delegate of is an error of the function,
                // not of the drawing (see CompileExpression)
                result.AddError(ex.InnerException?.Message ?? ex.Message);
            }
            finally
            {
                CompileTime += Stopwatch.GetElapsedTime(started);
            }

            return result;
        }

        public CompileResult CompileExpression(
            Drawing drawing,
            string expressionText,
            Predicate<IFigure> isFigureAllowed)
        {
            long started = Stopwatch.GetTimestamp();
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
                var bound = builder.CreateExpression(ast, result);
                if (bound == null || !result.Errors.IsEmpty())
                {
                    return result;
                }

                result.Expression = MakeExpression(bound);
            }
            catch (Exception ex)
            {
                // an expression the builder can't make sense of (a function of an expression
                // where it expects a point, say) is an error of the label, not of the drawing
                result.AddError(ex.InnerException?.Message ?? ex.Message);
            }
            finally
            {
                CompileTime += Stopwatch.GetElapsedTime(started);
            }

            return result;
        }

        Func<double, double> MakeFunction(BoundExpression bound)
        {
            if (Strategy == ExpressionStrategy.OwnTree)
            {
                return bound.ToFunctionDelegate();
            }

            var parameter = Expression.Parameter(typeof(double), "x");
            var lambda = Expression.Lambda<Func<double, double>>(bound.ToLinq(parameter), parameter);
            return Provider.InterpretFunction(lambda);
        }

        Func<double> MakeExpression(BoundExpression bound)
        {
            if (Strategy == ExpressionStrategy.OwnTree)
            {
                return bound.ToExpressionDelegate();
            }

            var lambda = Expression.Lambda<Func<double>>(bound.ToLinq(null));
            return Provider.InterpretExpression(lambda);
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
