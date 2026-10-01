using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace DynamicGeometry
{
    public class Binder
    {
        static Binder()
        {
            AddMethods(typeof(System.Math));
            AddMethods(typeof(Functions));
        }

        static void AddMethods(Type type)
        {
            foreach (var methodInfo in type.GetMethods())
            {
                methods.Add(methodInfo);
            }
        }

        static List<MethodInfo> methods = new List<MethodInfo>();

        public void RegisterParameter(ParameterExpression parameter)
        {
            parameters.Add(parameter.Name, parameter);
        }

        public Drawing Drawing { get; set; }
        public Predicate<IFigure> FigureAllowed { get; set; }

        ParameterExpression ResolveParameter(string parameterName)
        {
            ParameterExpression parameter;
            if (parameters.TryGetValue(parameterName, out parameter))
            {
                return parameter;
            }
            return null;
        }

        Expression ResolveConstant(string identifier)
        {
            if (identifier.Equals("pi", StringComparison.InvariantCultureIgnoreCase))
            {
                return Expression.Constant(Math.PI);
            }
            else if (identifier.Equals("e", StringComparison.InvariantCultureIgnoreCase))
            {
                return Expression.Constant(System.Math.E);
            }
            return null;
        }

        Dictionary<string, ParameterExpression> parameters = new Dictionary<string, ParameterExpression>();

        public Expression Resolve(string identifier)
        {
            return ResolveConstant(identifier) ?? ResolveParameter(identifier);
        }

        /// <summary>
        /// The function called by this name with this many arguments: one of System.Math
        /// that takes that many numbers (sin, sqrt, max, atan2) if there is one, else
        /// whatever goes by the name (ours take points: dist, ang, area)
        /// </summary>
        public MethodInfo ResolveMethod(string functionName, int argumentCount)
        {
            foreach (var methodInfo in typeof(System.Math).GetMethods())
            {
                if (methodInfo.Name.Equals(functionName, StringComparison.OrdinalIgnoreCase)
                    && TakesNumbers(methodInfo, argumentCount))
                {
                    return methodInfo;
                }
            }

            foreach (var methodInfo in methods)
            {
                if (methodInfo.Name.Equals(functionName, StringComparison.OrdinalIgnoreCase))
                {
                    return methodInfo;
                }
            }

            return null;
        }

        /// <summary>A function of numbers, called with as many as it takes: its arguments are expressions, not the names of points</summary>
        public static bool TakesNumbers(MethodInfo method, int argumentCount)
        {
            var parameters = method.GetParameters();
            return parameters.Length == argumentCount
                && argumentCount > 0
                && parameters.All(parameter => parameter.ParameterType == typeof(double));
        }

        /// <summary>
        /// A number of the drawing (a slider, a Number) with exactly this name, capitals
        /// and all; not pi, e or a function's x, which stay what they are. Null when there
        /// is none.
        /// </summary>
        public INumber ResolveExactNumber(string name)
        {
            if (Drawing == null || Resolve(name) != null)
            {
                return null;
            }

            return Drawing.Figures.FirstOrDefault(f => f is INumber && f.Name == name) as INumber;
        }

        public IFigure ResolveFigure(string figureName)
        {
            var candidate = Drawing.Figures[figureName];
            if (candidate == null)
            {
                // the indexer looks inside a composite figure (a slider) and not at it, and a
                // slider "a" must not be taken for the point "A" by the lenient search below
                candidate = Drawing.Figures.FirstOrDefault(f => f != null && f.Name == figureName);
            }
            if (candidate == null)
            {
                candidate = Drawing.Figures
                    .Where(f => f != null 
                        && !f.Name.IsEmpty() 
                        && f.Name.Equals(figureName, StringComparison.OrdinalIgnoreCase))
                    .FirstOrDefault();
            }
            if (candidate == null)
            {
                return null;
            }
            return candidate;
        }

        public bool IsFigureAllowed(IFigure candidate)
        {
            if (FigureAllowed == null)
            {
                return true;
            }
            return FigureAllowed(candidate);
        }
    }
}
