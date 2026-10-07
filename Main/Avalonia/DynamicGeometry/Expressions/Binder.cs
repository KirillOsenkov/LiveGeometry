using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DynamicGeometry
{
    /// <summary>
    /// Resolves the names of one expression against a drawing: figures, numbers, the x of a
    /// function, the points whose names run together into a distance, and the functions
    /// (through <see cref="ExpressionReflection"/>, which looks each up once for all). A
    /// binder binds one expression and remembers what its names stood for: an expression says
    /// the same name several times (A.X - B.X, A.Y - B.Y), and each lookup is a search of the
    /// whole drawing.
    /// </summary>
    public class Binder
    {
        string parameterName;

        /// <summary>The x of a function: a name that stands for the parameter, not for a figure</summary>
        public void RegisterParameter(string name)
        {
            parameterName = name;
        }

        public Drawing Drawing { get; set; }
        public Predicate<IFigure> FigureAllowed { get; set; }

        BoundExpression ResolveParameter(string identifier)
        {
            return parameterName != null && identifier == parameterName ? new BoundParameter() : null;
        }

        BoundExpression ResolveConstant(string identifier)
        {
            if (identifier.Equals("pi", StringComparison.InvariantCultureIgnoreCase))
            {
                return new BoundConstant(Math.PI);
            }
            else if (identifier.Equals("e", StringComparison.InvariantCultureIgnoreCase))
            {
                return new BoundConstant(System.Math.E);
            }
            return null;
        }

        /// <summary>pi, e or the x of a function; null for any other name</summary>
        public BoundExpression Resolve(string identifier)
        {
            return ResolveConstant(identifier) ?? ResolveParameter(identifier);
        }

        /// <summary>See <see cref="ExpressionReflection.ResolveMethod"/></summary>
        public MethodInfo ResolveMethod(string functionName, int argumentCount)
        {
            return ExpressionReflection.ResolveMethod(functionName, argumentCount);
        }

        /// <summary>A function of numbers, called with as many as it takes: its arguments are expressions, not the names of points</summary>
        public static bool TakesNumbers(MethodInfo method, int argumentCount)
        {
            return ExpressionReflection.TakesNumbers(method, argumentCount);
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

            if (!numbersByName.TryGetValue(name, out var number))
            {
                number = Drawing.Figures.FirstOrDefault(f => f is INumber && f.Name == name) as INumber;
                numbersByName[name] = number;
            }

            return number;
        }

        // What the names of the expression stand for, looked up once each (a binder binds one
        // expression): an expression says the same name several times (A.X - B.X, A.Y - B.Y),
        // and each lookup is a search of the whole drawing
        readonly Dictionary<string, INumber> numbersByName = new Dictionary<string, INumber>();
        readonly Dictionary<string, IFigure> figuresByName = new Dictionary<string, IFigure>();
        string[] pointNames;

        /// <summary>
        /// The names of the drawing's points, for reading two of them run together (AB):
        /// gathered once per expression, where each identifier gathered them anew
        /// </summary>
        public string[] PointNames
        {
            get
            {
                if (pointNames == null)
                {
                    pointNames = Drawing == null
                        ? Array.Empty<string>()
                        : Drawing.Figures.Where(f => f is PointBase).Select(f => f.Name).ToArray();
                }

                return pointNames;
            }
        }

        public IFigure ResolveFigure(string figureName)
        {
            if (!figuresByName.TryGetValue(figureName, out var figure))
            {
                figure = FindFigure(figureName);
                figuresByName[figureName] = figure;
            }

            return figure;
        }

        IFigure FindFigure(string figureName)
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
