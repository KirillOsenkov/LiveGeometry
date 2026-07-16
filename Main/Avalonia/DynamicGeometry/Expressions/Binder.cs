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

        public MethodInfo ResolveMethod(string functionName)
        {
            foreach (var methodInfo in typeof(System.Math).GetMethods())
            {
                var parameters = methodInfo.GetParameters();
                if (methodInfo.Name.Equals(functionName, StringComparison.OrdinalIgnoreCase)
                    && parameters.Length == 1
                    && parameters[0].ParameterType == typeof(double))
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

        public IFigure ResolveFigure(string figureName)
        {
            var candidate = Drawing.Figures[figureName];
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
