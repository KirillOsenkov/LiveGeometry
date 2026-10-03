using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace DynamicGeometry
{
    public class ExpressionTreeBuilder
    {
        public ExpressionTreeBuilder()
        {
            Binder = new Binder();
        }

        public Binder Binder { get; set; }
        CompileResult Status { get; set; }

        public Expression<Func<double, double>> CreateFunction(Node root, CompileResult status)
        {
            Status = status;
            ParameterExpression parameter = Expression.Parameter(typeof(double), "x");
            Binder.RegisterParameter(parameter);
            Expression body = CreateExpressionCore(root);
            if (body == null)
            {
                return null;
            }
            var expressionTree = Expression.Lambda<Func<double, double>>(body, parameter);
            return expressionTree;
        }

        public Expression<Func<double>> CreateExpression(Node root, CompileResult status)
        {
            Status = status;
            Expression body = CreateExpressionCore(root);
            if (body == null)
            {
                return null;
            }
            // If expression does not return double, we'll get an error. - D.H.
            if (body.Type != typeof(double))
            {
                return null;
            }
            var expressionTree = Expression.Lambda<Func<double>>(body);
            return expressionTree;
        }

        Expression CreateExpressionCore(Node root)
        {
            switch (root.Kind)
            {
                case NodeType.Negation:
                    return CreateUnaryExpression(root);
                case NodeType.Addition:
                case NodeType.Subtraction:
                case NodeType.Multiplication:
                case NodeType.Division:
                case NodeType.Power:
                    return CreateBinaryExpression(root);
                case NodeType.Variable:
                    return CreateIdentifierExpression(root);
                case NodeType.Constant:
                    // the language writes decimals with a point, whatever the user's culture
                    return CreateLiteralExpression(double.Parse(root.Token.Text, CultureInfo.InvariantCulture));
                case NodeType.FunctionCall:
                    return CreateCallExpression(root);
                case NodeType.PropertyAccess:
                    return CreatePropertyAccessExpression(root);
                default:
                    return null;
            }
        }

        Expression CreateUnaryExpression(Node root)
        {
            Expression operand = CreateExpressionCore(root.Children[0]);
            if (operand == null)
            {
                return null;
            }

            return Expression.Negate(operand);
        }

        Expression CreateIdentifierExpression(Node root)
        {
            var text = root.Token.Text;

            // A number called exactly this comes before the reading as two points, which
            // takes the points' names in any case: with points A and B in the drawing, a
            // slider named ab was the distance AB.
            var exact = Binder.ResolveExactNumber(text);
            if (exact != null)
            {
                return CreateNumberExpression(exact, text);
            }

            // pi is π whatever the drawing has. Two points are read in any case (ab is the
            // distance AB), so in a drawing with points P and I it was their distance, and
            // sin(pi * x) came out wrong without a word. Written as the points are, PI, it
            // still is the distance.
            if (text == "pi" || text == "e")
            {
                return Binder.Resolve(text);
            }

            Expression resolveTwoPoints = ResolveTwoPoints(text);
            if (resolveTwoPoints != null)
            {
                return resolveTwoPoints;
            }

            var parameter = Binder.Resolve(text);
            if (parameter != null)
            {
                return parameter;
            }

            // a Number in the drawing, by its name: the expression then depends on it
            if (Binder.ResolveFigure(text) is INumber number)
            {
                return CreateNumberExpression(number, text);
            }

            Status.AddUnknownIdentifierError(text);
            return null;
        }

        Expression CreateNumberExpression(INumber number, string text)
        {
            if (!Binder.IsFigureAllowed(number))
            {
                Status.AddDependencyCycleError(text);
                return null;
            }

            Status.Dependencies.Add(number);
            return Expression.Property(
                Expression.Constant(number, typeof(INumber)),
                typeof(INumber).GetProperty("Value"));
        }

        public Expression ResolveTwoPoints(string twoPoints)
        {
            var drawing = Binder.Drawing;
            if (drawing == null)
            {
                return null;
            }

            var names = drawing.Figures.Where(f => f is PointBase).Select(f => f.Name).ToArray();
            string longestPrefix = "";
            string longestSuffix = "";
            foreach (var name in names)
            {
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                if (twoPoints.StartsWith(name, StringComparison.OrdinalIgnoreCase) && name.Length > longestPrefix.Length)
                {
                    longestPrefix = name;
                }
                if (twoPoints.EndsWith(name, StringComparison.OrdinalIgnoreCase) && name.Length > longestSuffix.Length)
                {
                    longestSuffix = name;
                }
            }

            if (longestPrefix.Length + longestSuffix.Length == twoPoints.Length)
            {
                PointBase point1 = drawing.Figures[longestPrefix] as PointBase;
                PointBase point2 = drawing.Figures[longestSuffix] as PointBase;

                if (point1 == null)
                {
                    Status.AddFigureIsNotAPointError(longestPrefix);
                    return null;
                }
                if (point2 == null)
                {
                    Status.AddFigureIsNotAPointError(longestSuffix);
                    return null;
                }
                if (!Binder.IsFigureAllowed(point1))
                {
                    Status.AddDependencyCycleError(longestPrefix);
                    return null;
                }
                if (!Binder.IsFigureAllowed(point2))
                {
                    Status.AddDependencyCycleError(longestSuffix);
                    return null;
                }

                ConstantExpression p1 = Expression.Constant(point1);
                ConstantExpression p2 = Expression.Constant(point2);
                MethodInfo distance = typeof(Math).GetMethod("Distance",
                    new[] { typeof(PointBase), typeof(PointBase) });
                MethodCallExpression result = Expression.Call(null, distance, p1, p2);
                Status.Dependencies.Add(point1);
                Status.Dependencies.Add(point2);
                return result;
            }

            return null;
        }

        Expression CreatePropertyAccessExpression(Node root)
        {
            string figureName = root.Children[0].Token.Text;
            string propertyName = root.Children[1].Token.Text;

            IFigure figure = Binder.ResolveFigure(figureName);
            if (figure == null)
            {
                Status.AddUnknownIdentifierError(figureName);
                return null;
            }

            if (!Binder.IsFigureAllowed(figure))
            {
                Status.AddDependencyCycleError(figureName);
                return null;
            }

            var property = FindProperty(figure.GetType(), propertyName);
            if (property == null)
            {
                Status.AddPropertyNotFoundError(figure, propertyName);
                return null;
            }

            var figureExpression = Expression.Constant(figure);
            var value = AsNumber(Expression.Property(figureExpression, property), figureName + "." + propertyName);
            if (value != null)
            {
                Status.Dependencies.Add(figure);
            }

            return value;
        }

        // the property a name means on a type of figure, looked up once (A.X in one expression
        // after another went through all the properties of a point each time)
        static readonly ConcurrentDictionary<(Type Type, string Name), PropertyInfo> properties
            = new ConcurrentDictionary<(Type Type, string Name), PropertyInfo>();

        static PropertyInfo FindProperty(Type type, string propertyName)
        {
            return properties.GetOrAdd((type, propertyName), key =>
            {
                // by name in any case, the one written exactly first (GetProperty with
                // IgnoreCase throws when two differ only by case, or one hides another)
                var candidates = key.Type
                    .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(p => p.Name.Equals(key.Name, StringComparison.OrdinalIgnoreCase) && p.GetIndexParameters().Length == 0 && p.CanRead)
                    .ToList();
                return candidates.FirstOrDefault(p => p.Name == key.Name) ?? candidates.FirstOrDefault();
            });
        }

        Expression CreateCallExpression(Node root)
        {
            string functionName = root.Token.Text;
            var arguments = root.Children;
            MethodInfo method = Binder.ResolveMethod(functionName, arguments.Count);
            if (method == null)
            {
                Status.AddMethodNotFoundError(functionName);
                return null;
            }

            // a function of numbers (sin, sqrt, max, atan2): each argument is an expression.
            // Anything else takes points (dist, ang) - and a wrong number of arguments is an
            // error, not an exception out of Expression.Call. (max(x, 2) used to be told
            // that it "takes the names of points".)
            if (Binder.TakesNumbers(method, arguments.Count))
            {
                var values = new List<Expression>();
                foreach (var node in arguments)
                {
                    var value = node == null ? null : CreateExpressionCore(node);
                    if (value == null)
                    {
                        return null;
                    }

                    values.Add(value);
                }

                return AsNumber(Expression.Call(method, values), functionName);
            }

            return CreatePointFunctionCallExpression(method, arguments);
        }

        /// <summary>
        /// What the language calculates with is a double. A whole number becomes one (the
        /// sign of a number, a polygon's count of sides: left as it was, the first operator
        /// applied to it threw, and a function that ended in it could not be made). Anything
        /// else - a name, a check box, a point - is an error, said here.
        /// </summary>
        Expression AsNumber(Expression value, string what)
        {
            if (value.Type == typeof(double))
            {
                return value;
            }

            if (value.Type == typeof(int) || value.Type == typeof(long) || value.Type == typeof(float) || value.Type == typeof(decimal))
            {
                return Expression.Convert(value, typeof(double));
            }

            Status.AddError(string.Format("'{0}' is not a number", what));
            return null;
        }

        Expression CreatePointFunctionCallExpression(MethodInfo method, IEnumerable<Node> arguments)
        {
            if (method.Name != "Area" && method.GetParameters().Length != arguments.Count())
            {
                Status.AddIncorrectNumberOfArgumentsError(method, arguments.Count());
                return null;
            }

            List<IPoint> points = new List<IPoint>();
            foreach (var node in arguments)
            {
                // an argument that is an expression or a number, not a name (VB6 drawings have
                // those)
                if (node.Token == null || node.Token.Kind != TokenType.Identifier)
                {
                    Status.AddError(string.Format("'{0}' takes the names of points", method.Name));
                    return null;
                }

                string pointName = node.Token.Text;
                var point = ResolvePoint(pointName);
                if (point == null)
                {
                    return null;
                }
                points.Add(point);
            }

            if (method.Name == "Area")
            {
                return Expression.Call(method, Expression.Constant(points.ToArray()));
            }

            List<Expression> pointArguments = new List<Expression>();
            foreach (var point in points)
            {
                Expression pointArgument = CreatePointExpression(point);
                if (pointArgument == null)
                {
                    return null;
                }
                pointArguments.Add(pointArgument);
            }

            return Expression.Call(method, pointArguments.ToArray());
        }

        Expression CreatePointExpression(IPoint point)
        {
            var pointExpression = Expression.Constant(point);
            var coordinatesProperty = typeof(IPoint).GetProperty("Coordinates");
            var coordinates = Expression.Property(pointExpression, coordinatesProperty);
            return coordinates;
        }

        IPoint ResolvePoint(string pointName)
        {
            IFigure figure = Binder.ResolveFigure(pointName);
            if (figure == null)
            {
                Status.AddUnknownIdentifierError(pointName);
                return null;
            }
            IPoint point = figure as IPoint;
            if (point == null)
            {
                Status.AddFigureIsNotAPointError(pointName);
                return null;
            }

            if (!Binder.IsFigureAllowed(point))
            {
                Status.AddDependencyCycleError(pointName);
                return null;
            }

            Status.Dependencies.Add(point);
            return point;
        }

        Expression CreateLiteralExpression(double arg)
        {
            return Expression.Constant(arg);
        }

        Expression CreateBinaryExpression(Node node)
        {
            Expression left = CreateExpressionCore(node.Children[0]);
            Expression right = CreateExpressionCore(node.Children[1]);

            if (left == null || right == null)
            {
                return null;
            }

            switch (node.Kind)
            {
                case NodeType.Addition:
                    return Expression.Add(left, right);
                case NodeType.Subtraction:
                    return Expression.Subtract(left, right);
                case NodeType.Multiplication:
                    return Expression.Multiply(left, right);
                case NodeType.Division:
                    return Expression.Divide(left, right);
                case NodeType.Power:
                    return Expression.Power(left, right);
            }
            return null;
        }

        public void SetContext(Drawing drawing, Predicate<IFigure> isFigureAllowed)
        {
            Binder.Drawing = drawing;
            Binder.FigureAllowed = isFigureAllowed;
        }
    }
}
