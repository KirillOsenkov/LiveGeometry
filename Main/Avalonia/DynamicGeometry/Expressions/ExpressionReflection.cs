using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// What binding and evaluating expressions ask reflection for, looked up once: a function by
/// its name and the number of its arguments, the shape of a method, a property's getter and
/// a method as delegates. Before (2026-10-07) every compile went through System.Math's
/// methods for each function it named, and the interpreter of a System.Linq.Expressions tree
/// invoked methods and read properties by reflection at every evaluation.
/// </summary>
public static class ExpressionReflection
{
    public static readonly PropertyInfo NumberValue = typeof(INumber).GetProperty(nameof(INumber.Value));

    public static readonly PropertyInfo PointCoordinates = typeof(IPoint).GetProperty(nameof(IPoint.Coordinates));

    public static readonly MethodInfo DistanceOfPoints = typeof(Math).GetMethod(nameof(Math.Distance), new[] { typeof(PointBase), typeof(PointBase) });

    /// <summary>
    /// The functions of System.Math are found by name, through reflection the trimmer
    /// can't follow (a type out of a list): it kept only those the app calls itself, and
    /// in the browser build asinh, sinh, cosh... were "Could not find method" - the
    /// Catenary drew no curve. Kept whole here.
    /// </summary>
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicMethods, typeof(System.Math))]
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicMethods, typeof(Functions))]
    static ExpressionReflection()
    {
        // ours first: round, sign and clamp are System.Math's names, done as a drawing wants them
        methods = typeof(Functions).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Concat(typeof(System.Math).GetMethods(BindingFlags.Public | BindingFlags.Static))
            .ToArray();
        KeepGetterInstantiations();
    }

    // Mono's AOT compiles the generic instantiations it sees in the code, and runs one it
    // hasn't compiled in its interpreter. These are the ones MakeGetter asks for (every owner
    // is a class, and the instantiation over object serves them all): named here, the getters
    // are native code in the browser too. Nothing is called.
    static void KeepGetterInstantiations()
    {
        Func<MethodInfo, Func<object, double>>[] instantiations =
        {
            WrapGetter<object, double>,
            WrapGetter<object, int>,
            WrapGetter<object, long>,
            WrapGetter<object, float>,
            WrapGetter<object, decimal>
        };
        GC.KeepAlive(instantiations);
    }

    static readonly MethodInfo[] methods;
    static readonly ConcurrentDictionary<(string Name, int Count), MethodInfo> resolvedMethods = new ConcurrentDictionary<(string Name, int Count), MethodInfo>();
    static readonly ConcurrentDictionary<MethodInfo, MethodShape> shapes = new ConcurrentDictionary<MethodInfo, MethodShape>();
    static readonly ConcurrentDictionary<MethodInfo, NumberFunction> numberFunctions = new ConcurrentDictionary<MethodInfo, NumberFunction>();
    static readonly ConcurrentDictionary<MethodInfo, PointFunction> pointFunctions = new ConcurrentDictionary<MethodInfo, PointFunction>();
    static readonly ConcurrentDictionary<PropertyInfo, Func<object, double>> getters = new ConcurrentDictionary<PropertyInfo, Func<object, double>>();

    /// <summary>
    /// The function called by this name with this many arguments: one that takes that
    /// many numbers if there is one - ours first (round, sign: the ones System.Math has
    /// in a way a drawing doesn't want), then System.Math's (sin, sqrt, max, atan2) -
    /// else whatever goes by the name (ours that take points: dist, ang, area). Null when
    /// there is none.
    /// </summary>
    public static MethodInfo ResolveMethod(string functionName, int argumentCount)
    {
        return resolvedMethods.GetOrAdd((functionName.ToLowerInvariant(), argumentCount), key => FindMethod(key.Name, key.Count));
    }

    static MethodInfo FindMethod(string functionName, int argumentCount)
    {
        foreach (var method in methods)
        {
            if (method.Name.Equals(functionName, StringComparison.OrdinalIgnoreCase) && TakesNumbers(method, argumentCount))
            {
                return method;
            }
        }

        foreach (var method in methods)
        {
            if (method.Name.Equals(functionName, StringComparison.OrdinalIgnoreCase))
            {
                return method;
            }
        }

        return null;
    }

    /// <summary>A function of numbers, called with as many as it takes: its arguments are expressions, not the names of points</summary>
    public static bool TakesNumbers(MethodInfo method, int argumentCount)
    {
        var shape = shapes.GetOrAdd(method, m => new MethodShape(m));
        return shape.ParameterCount == argumentCount && argumentCount > 0 && shape.AllDouble;
    }

    class MethodShape
    {
        public MethodShape(MethodInfo method)
        {
            var parameters = method.GetParameters();
            ParameterCount = parameters.Length;
            AllDouble = parameters.All(parameter => parameter.ParameterType == typeof(double));
        }

        public int ParameterCount { get; }

        public bool AllDouble { get; }
    }

    public static NumberFunction GetNumberFunction(MethodInfo method)
    {
        return numberFunctions.GetOrAdd(method, m => new NumberFunction(m));
    }

    public static PointFunction GetPointFunction(MethodInfo method)
    {
        return pointFunctions.GetOrAdd(method, m => new PointFunction(m));
    }

    /// <summary>The value of a figure's property as a double (whole numbers converted), read through a delegate rather than reflection</summary>
    public static Func<object, double> GetNumberGetter(PropertyInfo property)
    {
        return getters.GetOrAdd(property, MakeGetter);
    }

    /// <summary>What the language calculates with, or converts on the way (see <see cref="ExpressionTreeBuilder.AsNumber"/>)</summary>
    public static bool IsNumberType(Type type)
    {
        return type == typeof(double) || type == typeof(int) || type == typeof(long) || type == typeof(float) || type == typeof(decimal);
    }

    public static double ToDouble(object value)
    {
        return Convert.ToDouble(value, CultureInfo.InvariantCulture);
    }

    /// <summary>A Linq tree's value as a double: a whole number converted (the bound tree checked that it is a number)</summary>
    public static Expression AsDouble(Expression value)
    {
        return value.Type == typeof(double) ? value : Expression.Convert(value, typeof(double));
    }

    static Func<object, double> MakeGetter(PropertyInfo property)
    {
        var getter = property.GetGetMethod();
        var owner = getter?.DeclaringType;
        if (getter == null || getter.IsStatic || owner == null || owner.IsValueType || !IsNumberType(property.PropertyType))
        {
            return figure => ToDouble(property.GetValue(figure));
        }

        // an open delegate on the getter, Func<owner, value>, through a generic method so that
        // no code is written per property; a figure is a class, so the owner can stand in for all
        try
        {
            var wrap = typeof(ExpressionReflection)
                .GetMethod(nameof(WrapGetter), BindingFlags.NonPublic | BindingFlags.Static)
                .MakeGenericMethod(owner, property.PropertyType);
            return (Func<object, double>)wrap.Invoke(null, new object[] { getter });
        }
        catch (Exception ex) when (ex is ArgumentException || ex is TargetInvocationException || ex is NotSupportedException)
        {
            return figure => ToDouble(property.GetValue(figure));
        }
    }

    static Func<object, double> WrapGetter<TOwner, TValue>(MethodInfo getter)
        where TOwner : class
    {
        var typed = (Func<TOwner, TValue>)getter.CreateDelegate(typeof(Func<TOwner, TValue>));
        if (typed is Func<TOwner, double> exact)
        {
            return owner => exact((TOwner)owner);
        }

        return owner => ToDouble(typed((TOwner)owner));
    }

    /// <summary>
    /// A static method called with evaluated arguments: through its delegate where its shape
    /// is a function of numbers or of points (<see cref="NumberFunction"/>,
    /// <see cref="PointFunction"/>), else through reflection
    /// </summary>
    public static object Invoke(MethodInfo method, object[] arguments)
    {
        if (method == DistanceOfPoints && arguments.Length == 2)
        {
            return Math.Distance((PointBase)arguments[0], (PointBase)arguments[1]);
        }

        var shape = shapes.GetOrAdd(method, m => new MethodShape(m));
        if (shape.AllDouble && shape.ParameterCount == arguments.Length && arguments.Length > 0 && arguments.Length <= 3)
        {
            var function = GetNumberFunction(method);
            switch (arguments.Length)
            {
                case 1:
                    return function.Of1(ToDouble(arguments[0]));
                case 2:
                    return function.Of2(ToDouble(arguments[0]), ToDouble(arguments[1]));
                default:
                    return function.Of3(ToDouble(arguments[0]), ToDouble(arguments[1]), ToDouble(arguments[2]));
            }
        }

        if (arguments.Length > 0 && arguments.Length <= 3 && arguments.All(argument => argument is Point))
        {
            return GetPointFunction(method).InvokeCoordinates(arguments);
        }

        return method.Invoke(null, arguments);
    }
}

/// <summary>
/// A function of numbers (one of <see cref="Functions"/> or System.Math's) as delegates by
/// its number of arguments: a method that takes doubles and gives one is called directly;
/// one that gives a whole number (System.Math.ILogB) is invoked and its result converted.
/// </summary>
public class NumberFunction
{
    public NumberFunction(MethodInfo method)
    {
        Method = method;
        var parameters = method.GetParameters();
        bool exact = method.ReturnType == typeof(double) && parameters.All(parameter => parameter.ParameterType == typeof(double));
        Of1 = exact && parameters.Length == 1
            ? (Func<double, double>)method.CreateDelegate(typeof(Func<double, double>))
            : first => Invoke(new object[] { first });
        Of2 = exact && parameters.Length == 2
            ? (Func<double, double, double>)method.CreateDelegate(typeof(Func<double, double, double>))
            : (first, second) => Invoke(new object[] { first, second });
        Of3 = exact && parameters.Length == 3
            ? (Func<double, double, double, double>)method.CreateDelegate(typeof(Func<double, double, double, double>))
            : (first, second, third) => Invoke(new object[] { first, second, third });
        OfMany = values => Invoke(values.Cast<object>().ToArray());
    }

    public MethodInfo Method { get; }

    public Func<double, double> Of1 { get; }

    public Func<double, double, double> Of2 { get; }

    public Func<double, double, double, double> Of3 { get; }

    public Func<double[], double> OfMany { get; }

    double Invoke(object[] arguments)
    {
        return ExpressionReflection.ToDouble(Method.Invoke(null, arguments));
    }
}

/// <summary>
/// A function of points (dist, ang, area: one of <see cref="Functions"/>) as a delegate,
/// given the points it is called with (their coordinates as they are now)
/// </summary>
public class PointFunction
{
    readonly Func<IPoint[], double> ofArray;
    readonly Func<Point, double> of1;
    readonly Func<Point, Point, double> of2;
    readonly Func<Point, Point, Point, double> of3;

    public PointFunction(MethodInfo method)
    {
        Method = method;
        var parameters = method.GetParameters();
        TakesArray = parameters.Length == 1 && parameters[0].ParameterType == typeof(IPoint[]);
        bool exact = method.ReturnType == typeof(double) && parameters.All(parameter => parameter.ParameterType == typeof(Point));
        if (TakesArray && method.ReturnType == typeof(double))
        {
            ofArray = (Func<IPoint[], double>)method.CreateDelegate(typeof(Func<IPoint[], double>));
        }

        if (exact && parameters.Length == 1)
        {
            of1 = (Func<Point, double>)method.CreateDelegate(typeof(Func<Point, double>));
        }
        else if (exact && parameters.Length == 2)
        {
            of2 = (Func<Point, Point, double>)method.CreateDelegate(typeof(Func<Point, Point, double>));
        }
        else if (exact && parameters.Length == 3)
        {
            of3 = (Func<Point, Point, Point, double>)method.CreateDelegate(typeof(Func<Point, Point, Point, double>));
        }
    }

    public MethodInfo Method { get; }

    /// <summary>The function takes the points themselves, as many as there are (area)</summary>
    public bool TakesArray { get; }

    public double Invoke(IPoint[] points)
    {
        if (TakesArray)
        {
            return ofArray != null ? ofArray(points) : ExpressionReflection.ToDouble(Method.Invoke(null, new object[] { points }));
        }

        if (of1 != null && points.Length == 1)
        {
            return of1(points[0].Coordinates);
        }

        if (of2 != null && points.Length == 2)
        {
            return of2(points[0].Coordinates, points[1].Coordinates);
        }

        if (of3 != null && points.Length == 3)
        {
            return of3(points[0].Coordinates, points[1].Coordinates, points[2].Coordinates);
        }

        var coordinates = new object[points.Length];
        for (int i = 0; i < points.Length; i++)
        {
            coordinates[i] = points[i].Coordinates;
        }

        return ExpressionReflection.ToDouble(Method.Invoke(null, coordinates));
    }

    /// <summary>Called with the points' coordinates (boxed), as a tree walked by <see cref="ExpressionTreeInterpreter"/> has them</summary>
    public double InvokeCoordinates(object[] coordinates)
    {
        if (of1 != null && coordinates.Length == 1)
        {
            return of1((Point)coordinates[0]);
        }

        if (of2 != null && coordinates.Length == 2)
        {
            return of2((Point)coordinates[0], (Point)coordinates[1]);
        }

        if (of3 != null && coordinates.Length == 3)
        {
            return of3((Point)coordinates[0], (Point)coordinates[1], (Point)coordinates[2]);
        }

        return ExpressionReflection.ToDouble(Method.Invoke(null, coordinates));
    }
}
