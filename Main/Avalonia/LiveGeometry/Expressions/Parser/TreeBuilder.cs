using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace LiveGeometry;

/// <summary>
/// Binds the parser's tree (<see cref="Node"/>) against a drawing into a
/// <see cref="BoundExpression"/>: every name a figure, a number, a distance or the x of a
/// function, every call a method, every value a number. What the bound tree then becomes
/// is the compiler's choice (<see cref="ExpressionStrategy"/>). An error is said in the
/// <see cref="CompileResult"/> and the tree is null; nothing a user can type may throw.
/// </summary>
public class ExpressionTreeBuilder
{
    public ExpressionTreeBuilder()
    {
        Binder = new Binder();
    }

    public Binder Binder { get; set; }
    CompileResult Status { get; set; }

    public BoundExpression CreateFunction(Node root, CompileResult status)
    {
        Status = status;
        Binder.RegisterParameter("x");
        return CreateExpressionCore(root);
    }

    public BoundExpression CreateExpression(Node root, CompileResult status)
    {
        Status = status;
        return CreateExpressionCore(root);
    }

    BoundExpression CreateExpressionCore(Node root)
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
                return Classified(new BoundConstant(double.Parse(root.Token.Text, CultureInfo.InvariantCulture)), root.Token, ExpressionClass.Number);
            case NodeType.FunctionCall:
                return CreateCallExpression(root);
            case NodeType.PropertyAccess:
                return CreatePropertyAccessExpression(root);
            case NodeType.Degrees:
                return CreateDegreesExpression(root);
            case NodeType.AbsoluteValue:
                return CreateAbsoluteValueExpression(root);
            case var comparison when Parser.IsComparison(comparison):
                return CreateComparisonExpression(root);
            default:
                return null;
        }
    }

    /// <summary>30°, a°: the number in radians, what trigonometry takes (GeoGebra's degree sign)</summary>
    BoundExpression CreateDegreesExpression(Node root)
    {
        var operand = root.Children[0] == null ? null : CreateExpressionCore(root.Children[0]);
        return operand == null ? null : Classified(new BoundBinary(BoundOperator.Multiply, operand, new BoundConstant(Math.PI / 180)), root.Token, ExpressionClass.Reserved);
    }

    /// <summary>|x|, as abs(x)</summary>
    BoundExpression CreateAbsoluteValueExpression(Node root)
    {
        var operand = root.Children[0] == null ? null : CreateExpressionCore(root.Children[0]);
        return operand == null ? null : new BoundCall(ExpressionReflection.Abs, [operand]);
    }

    /// <summary>a &lt; b, AB = CD: 1 where it holds and 0 where it doesn't, for if(...)</summary>
    BoundExpression CreateComparisonExpression(Node root)
    {
        var left = root.Children[0] == null ? null : CreateExpressionCore(root.Children[0]);
        var right = left == null || root.Children[1] == null ? null : CreateExpressionCore(root.Children[1]);
        return right == null ? null : new BoundCall(ExpressionReflection.Comparison(root.Kind), [left, right]);
    }

    BoundExpression CreateUnaryExpression(Node root)
    {
        BoundExpression operand = CreateExpressionCore(root.Children[0]);
        if (operand == null)
        {
            return null;
        }

        return new BoundNegation(operand);
    }

    BoundExpression CreateIdentifierExpression(Node root)
    {
        var text = root.Token.Text;

        // a name in backticks is the number called exactly that: never pi, e, a
        // function's x or two points run together
        var token = root.Token;
        if (token.IsQuoted)
        {
            if (Binder.ResolveFigureExactly(text) is INumber quoted)
            {
                return Classified(CreateNumberExpression(quoted, text), token, ExpressionClass.Figure);
            }

            Status.AddUnknownIdentifierError(text);
            return null;
        }

        // A number called exactly this comes before the reading as two points, which
        // takes the points' names in any case: with points A and B in the drawing, a
        // slider named ab was the distance AB.
        var exact = Binder.ResolveExactNumber(text);
        if (exact != null)
        {
            return Classified(CreateNumberExpression(exact, text), token, ExpressionClass.Figure);
        }

        // pi is π whatever the drawing has. Two points are read in any case (ab is the
        // distance AB), so in a drawing with points P and I it was their distance, and
        // sin(pi * x) came out wrong without a word. Written as the points are, PI, it
        // still is the distance.
        if (text == "pi" || text == "π" || text == "e")
        {
            return Classified(Binder.Resolve(text), token, ExpressionClass.Reserved);
        }

        BoundExpression resolveTwoPoints = ResolveTwoPoints(text);
        if (resolveTwoPoints != null)
        {
            return Classified(resolveTwoPoints, token, ExpressionClass.Figure);
        }

        var parameter = Binder.Resolve(text);
        if (parameter != null)
        {
            return Classified(parameter, token, ExpressionClass.Reserved);
        }

        // a Number in the drawing, by its name: the expression then depends on it
        if (Binder.ResolveFigure(text) is INumber number)
        {
            return Classified(CreateNumberExpression(number, text), token, ExpressionClass.Figure);
        }

        Status.AddUnknownIdentifierError(text);
        return null;
    }

    /// <summary>
    /// The name or number bound, its piece of the text noted as what it was read as
    /// (<see cref="CompileResult.Classification"/>); one that didn't bind is noted as nothing
    /// </summary>
    BoundExpression Classified(BoundExpression bound, Token token, ExpressionClass kind)
    {
        if (bound != null)
        {
            Status.Classify(token, kind);
        }

        return bound;
    }

    BoundExpression CreateNumberExpression(INumber number, string text)
    {
        if (!Binder.IsFigureAllowed(number))
        {
            Status.AddDependencyCycleError(text);
            return null;
        }

        Status.Dependencies.Add(number);
        return new BoundNumber(number);
    }

    public BoundExpression ResolveTwoPoints(string twoPoints)
    {
        var drawing = Binder.Drawing;
        if (drawing == null)
        {
            return null;
        }

        string longestPrefix = "";
        string longestSuffix = "";
        foreach (var name in Binder.PointNames)
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

            Status.Dependencies.Add(point1);
            Status.Dependencies.Add(point2);
            return new BoundDistance(point1, point2);
        }

        return null;
    }

    BoundExpression CreatePropertyAccessExpression(Node root)
    {
        var figureToken = root.Children[0].Token;
        string figureName = figureToken.Text;
        string propertyName = root.Children[1].Token.Text;

        IFigure figure = ResolveFigure(figureToken);
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

        if (!AsNumber(property.PropertyType, figureName + "." + propertyName))
        {
            return null;
        }

        Status.Dependencies.Add(figure);
        Status.Classify(figureToken, ExpressionClass.Figure);
        Status.Classify(root.Children[1].Token, ExpressionClass.Property);
        return new BoundProperty(figure, property);
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

    BoundExpression CreateCallExpression(Node root)
    {
        string functionName = root.Token.Text;
        var arguments = root.Children;
        if (ExpressionReflection.LogBaseOf(functionName) is double logBase)
        {
            return CreateLogBaseExpression(root, logBase);
        }

        var function = FunctionCatalog.Find(functionName);
        if (function == null)
        {
            Status.AddMethodNotFoundError(functionName);
            return null;
        }

        // ∠ABC, angle(ABC): the points in one name, run together as the two of a distance are (AB)
        int joinedCount = JoinedPointCount(root);
        if (joinedCount > 0)
        {
            Status.Classify(root.Token, ExpressionClass.Function);
            return CreateJoinedPointsCallExpression(function, joinedCount, arguments[0].Token);
        }

        // angle(∠ABC) is ∠ABC
        if (function.Name == "angle"
            && arguments.Count == 1
            && arguments[0] is { Kind: NodeType.FunctionCall } inner
            && inner.Token.Kind == TokenType.Angle)
        {
            Status.Classify(root.Token, ExpressionClass.Function);
            return CreateExpressionCore(inner);
        }

        MethodInfo method = Binder.ResolveMethod(functionName, arguments.Count);
        if (method == null)
        {
            Status.AddIncorrectNumberOfArgumentsError(function, arguments.Count);
            return null;
        }

        // (before its arguments, which come after it in the text)
        Status.Classify(root.Token, ExpressionClass.Function);

        // a function of numbers (sin, sqrt, max, atan2): each argument is an expression.
        // Anything else takes points (distance, angle) - and a wrong number of arguments
        // is an error, not an exception out of Expression.Call.
        if (ExpressionReflection.IsNumberFunction(method))
        {
            if (FunctionCatalog.TakesRadians(function) && DegreesArgument(arguments) is string degrees)
            {
                Status.AddError(string.Format("{0} works in radians, and {1} gives degrees: put ° after it, {1}(...)°.", function.Name, degrees));
                return null;
            }

            var values = new List<BoundExpression>();
            foreach (var node in arguments)
            {
                var value = node == null ? null : CreateExpressionCore(node);
                if (value == null)
                {
                    return null;
                }

                values.Add(value);
            }

            if (!AsNumber(method.ReturnType, functionName))
            {
                return null;
            }

            // max(a, b, c) is max(max(a, b), c)
            BoundExpression call = new BoundCall(method, values.Take(method.GetParameters().Length).ToList());
            foreach (var value in values.Skip(method.GetParameters().Length))
            {
                call = new BoundCall(method, [call, value]);
            }

            return call;
        }

        return CreatePointFunctionCallExpression(function, method, arguments);
    }

    /// <summary>
    /// The name of a function of points that gives degrees (angle, directedangle, direction)
    /// when an argument is a call of it as it stands - sin(angle(A, B, C)), where the sine
    /// takes radians: the mistake the units of the language invite. Null when there is none.
    /// </summary>
    static string DegreesArgument(IEnumerable<Node> arguments)
    {
        foreach (var argument in arguments)
        {
            var inner = argument;
            while (inner != null && inner.Kind == NodeType.Negation)
            {
                inner = inner.Children[0];
            }

            if (inner != null && inner.Kind == NodeType.FunctionCall && FunctionCatalog.Find(inner.Token.Text) is FunctionInfo called && FunctionCatalog.GivesDegrees(called))
            {
                return called.Name;
            }
        }

        return null;
    }

    /// <summary>log_2(x): the logarithm to the base its name says</summary>
    BoundExpression CreateLogBaseExpression(Node root, double logBase)
    {
        if (root.Children.Count != 1)
        {
            Status.AddError(string.Format("'{0}' takes one number: {0}(x)", root.Token.Text));
            return null;
        }

        Status.Classify(root.Token, ExpressionClass.Function);
        var value = root.Children[0] == null ? null : CreateExpressionCore(root.Children[0]);
        return value == null ? null : new BoundCall(ExpressionReflection.LogBase, [value, new BoundConstant(logBase)]);
    }

    /// <summary>
    /// What the language calculates with is a double. A whole number becomes one (the
    /// sign of a number, a polygon's count of sides: left as it was, the first operator
    /// applied to it threw, and a function that ended in it could not be made). Anything
    /// else - a name, a check box, a point - is an error, said here.
    /// </summary>
    bool AsNumber(Type type, string what)
    {
        if (ExpressionReflection.IsNumberType(type))
        {
            return true;
        }

        Status.AddError(string.Format("'{0}' is not a number", what));
        return false;
    }

    BoundExpression CreatePointFunctionCallExpression(FunctionInfo function, MethodInfo method, IEnumerable<Node> arguments)
    {
        List<IPoint> points = new List<IPoint>();
        foreach (var node in arguments)
        {
            // an argument that is an expression or a number, not a name (VB6 drawings have
            // those)
            if (node == null || node.Kind != NodeType.Variable || node.Token == null || node.Token.Kind != TokenType.Identifier)
            {
                Status.AddError(string.Format("'{0}' takes the names of points: {1}", function.Name, function.Signature));
                return null;
            }

            var point = ResolvePoint(node.Token);
            if (point == null)
            {
                return null;
            }
            points.Add(point);
        }

        return AsNumber(method.ReturnType, method.Name) ? new BoundPointCall(method, points) : null;
    }

    /// <summary>
    /// How many points a call names in its one argument, their names run together: 3 for
    /// ∠ABC and angle(ABC), 2 for distance(AB). 0 for any other call (angle(A, B, C), sin(x),
    /// a name in backticks, which is one figure's).
    /// </summary>
    public static int JoinedPointCount(Node call)
    {
        if (call.Kind != NodeType.FunctionCall
            || call.Children.Count != 1
            || call.Children[0] is not { Kind: NodeType.Variable } argument
            || argument.Token == null
            || argument.Token.IsQuoted)
        {
            return 0;
        }

        var function = FunctionCatalog.Find(call.Token.Text);
        if (function == null || function.IsVariadic || function.ArgumentCount < 2)
        {
            return 0;
        }

        return ExpressionReflection.TakesPoints(function.Name, function.ArgumentCount) ? function.ArgumentCount : 0;
    }

    /// <summary>
    /// The text read as the names of <paramref name="count"/> of the items run together, in
    /// order (ABC: A, B and C): as written first, then in any case, as two points are read;
    /// a longer name before a shorter one. Null when it doesn't read so.
    /// </summary>
    public static List<T> SplitNames<T>(string text, int count, IEnumerable<T> items, Func<T, string> nameOf)
    {
        var named = items
            .Where(item => !string.IsNullOrEmpty(nameOf(item)))
            .OrderByDescending(item => nameOf(item).Length)
            .ToList();
        foreach (var comparison in new[] { StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase })
        {
            var found = new List<T>();
            if (Split(0))
            {
                return found;
            }

            bool Split(int start)
            {
                if (found.Count == count)
                {
                    return start == text.Length;
                }

                foreach (var item in named)
                {
                    var name = nameOf(item);
                    if (start + name.Length <= text.Length && string.Compare(text, start, name, 0, name.Length, comparison) == 0)
                    {
                        found.Add(item);
                        if (Split(start + name.Length))
                        {
                            return true;
                        }

                        found.RemoveAt(found.Count - 1);
                    }
                }

                return false;
            }
        }

        return null;
    }

    /// <summary>∠ABC, angle(ABC): a function of points called with their names run together into one</summary>
    BoundExpression CreateJoinedPointsCallExpression(FunctionInfo function, int count, Token joined)
    {
        var names = SplitNames(joined.Text, count, Binder.PointNames, name => name);
        if (names == null)
        {
            Status.AddError(string.Format("'{0}' is not the names of {1} points run together: {2}", joined.Text, count, function.Signature));
            return null;
        }

        var points = new List<IPoint>();
        foreach (var name in names)
        {
            var point = ResolvePoint(Binder.ResolveFigure(name), name);
            if (point == null)
            {
                return null;
            }

            points.Add(point);
        }

        Status.Classify(joined, ExpressionClass.Figure);
        var method = Binder.ResolveMethod(function.Name, count);
        return AsNumber(method.ReturnType, method.Name) ? new BoundPointCall(method, points) : null;
    }

    /// <summary>The figure a name stands for: in backticks, the one called exactly that</summary>
    IFigure ResolveFigure(Token token)
    {
        return token.IsQuoted ? Binder.ResolveFigureExactly(token.Text) : Binder.ResolveFigure(token.Text);
    }

    IPoint ResolvePoint(Token token)
    {
        var point = ResolvePoint(ResolveFigure(token), token.Text);
        if (point != null)
        {
            Status.Classify(token, ExpressionClass.Figure);
        }

        return point;
    }

    IPoint ResolvePoint(IFigure figure, string pointName)
    {
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

    BoundExpression CreateBinaryExpression(Node node)
    {
        BoundExpression left = CreateExpressionCore(node.Children[0]);
        BoundExpression right = CreateExpressionCore(node.Children[1]);

        if (left == null || right == null)
        {
            return null;
        }

        switch (node.Kind)
        {
            case NodeType.Addition:
                return new BoundBinary(BoundOperator.Add, left, right);
            case NodeType.Subtraction:
                return new BoundBinary(BoundOperator.Subtract, left, right);
            case NodeType.Multiplication:
                return new BoundBinary(BoundOperator.Multiply, left, right);
            case NodeType.Division:
                return new BoundBinary(BoundOperator.Divide, left, right);
            case NodeType.Power:
                return new BoundBinary(BoundOperator.Power, left, right);
        }
        return null;
    }

    public void SetContext(Drawing drawing, Predicate<IFigure> isFigureAllowed)
    {
        Binder.Drawing = drawing;
        Binder.FigureAllowed = isFigureAllowed;
    }
}
