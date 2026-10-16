using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// What a box an expression is typed into is for: a point clicked into the X of a point by
/// coordinates goes in as A.X, into the text of a label in brackets
/// </summary>
public enum InsertContext
{
    /// <summary>A number: a formula, a radius, a slope</summary>
    Expression,

    /// <summary>The X of a point by coordinates: a point goes in as its X</summary>
    PointX,

    /// <summary>The Y of a point by coordinates: a point goes in as its Y</summary>
    PointY,

    /// <summary>The text of a label: a reference goes in as a [...] part unless the caret is in one already</summary>
    Text
}

/// <summary>
/// The box that has the keyboard while an expression is typed (<see cref="ExpressionTextBox"/>):
/// a click on a figure, or on its row in the Figure List, puts a reference to it there instead
/// of going to the tool (<see cref="ReferencePicker"/>). The drawing holds the one at hand
/// (<see cref="Drawing.InsertTarget"/>).
/// </summary>
public interface IExpressionInsertTarget
{
    Drawing Drawing { get; }

    /// <summary>The figure whose expression is typed, which nothing it depends on may name; null for a tool's panel</summary>
    IFigure Owner { get; }

    InsertContext Context { get; }

    /// <summary>Puts the text in at the caret, over the selection if there is one</summary>
    void Insert(string reference);
}

/// <summary>One thing a click on a figure can put into the box: its text, and what that is in words</summary>
public class ReferenceOption
{
    public ReferenceOption(IFigure figure, string text, string description)
    {
        Figure = figure;
        Text = text;
        Description = description;
    }

    public IFigure Figure { get; }

    /// <summary>What goes into the box: "d.ArcLength", "AB", "angle(A, B, C)"</summary>
    public string Text { get; }

    /// <summary>"the arc length of arc d", for the status</summary>
    public string Description { get; }

    public bool SameAs(ReferenceOption other)
    {
        return other != null && Figure == other.Figure && Text == other.Text;
    }
}

/// <summary>
/// The Excel way of building a formula: while an expression box has the keyboard, a click on a
/// figure puts a reference to it at the caret (<see cref="Behavior"/> asks here before it
/// hands a press to the tool), and so does a click on a row of the Figure List, the way to a
/// hidden figure. A figure offers its first quantity as the Length tool would pick it, and the
/// others for Tab (<see cref="ReferenceText"/>, through <see cref="ClickChoice"/>): a circle
/// its radius, then its diameter, circumference and area; a measurement what it measures.
/// </summary>
public static class ReferencePicker
{
    /// <summary>What the status says while a box takes the keyboard</summary>
    public const string Hint = "Click a figure, or its row in the Figure List, to put it into the formula.";

    /// <summary>What a click here could put into the box, the figures as hit testing orders them, each with its quantities</summary>
    public static IReadOnlyList<ReferenceOption> FindOptions(Drawing drawing, Point coordinates, IExpressionInsertTarget target)
    {
        var options = new List<ReferenceOption>();
        var seen = new HashSet<IFigure>();
        foreach (var hit in drawing.Figures.HitTestAll(coordinates, figure => figure.Visible && figure.IsHitTestVisible))
        {
            var figure = Resolve(drawing, hit);
            if (figure != null && seen.Add(figure))
            {
                options.AddRange(OptionsFor(figure, target));
            }
        }

        return options;
    }

    /// <summary>
    /// Puts the figure's first quantity into the box (a row of the Figure List clicked); false,
    /// said in the status, when the figure has nothing to put there
    /// </summary>
    public static bool TryInsert(IExpressionInsertTarget target, IFigure figure)
    {
        var drawing = target.Drawing;
        var resolved = drawing != null ? Resolve(drawing, figure) : null;
        var option = resolved != null ? OptionsFor(resolved, target).FirstOrDefault() : null;
        if (option == null)
        {
            drawing?.RaiseStatusNotification("Nothing of " + ConstructionText.Of(figure) + " can go into the formula.");
            return false;
        }

        target.Insert(option.Text);
        return true;
    }

    /// <summary>The figure a click on this stands for: a name's figure, a part's whole; null for an axis</summary>
    static IFigure Resolve(Drawing drawing, IFigure hit)
    {
        if (hit is PointLabel || hit is FigureLabel)
        {
            hit = hit.Dependencies.FirstOrDefault();
        }

        if (hit == null || hit is AxisLine || hit is CartesianGrid)
        {
            return null;
        }

        return drawing.Figures.FindTopLevel(hit) ?? hit;
    }

    /// <summary>The figure's quantities as texts for this box; none for a figure the box's owner is built on the other way round</summary>
    public static IReadOnlyList<ReferenceOption> OptionsFor(IFigure figure, IExpressionInsertTarget target)
    {
        var owner = target.Owner;
        if (owner != null && (figure == owner || figure.DependsOn(owner)))
        {
            return [];
        }

        return ReferenceText.OptionsFor(figure, target.Context);
    }
}

/// <summary>
/// What an expression says for a figure's quantities: the figure's name in front of the
/// property the language reads (c.Radius, p.Area, u.X), two points run together for a
/// distance (AB, or dist(A, B) where the names won't run together), the functions of points
/// for an angle (angle(A, B, C) as the measurement's Sweep has it) and for the area of a list
/// of points. A measurement is said as what it measures, so that the formula reads like math
/// and outlives the measurement. The first option is what a click puts in.
/// </summary>
public static class ReferenceText
{
    public static IReadOnlyList<ReferenceOption> OptionsFor(IFigure figure, InsertContext context)
    {
        var options = new List<ReferenceOption>();
        void Add(string text, string description)
        {
            if (text != null)
            {
                options.Add(new ReferenceOption(figure, text, description));
            }
        }

        string name = Scanner.NameInExpression(figure.Name);
        string of = ConstructionText.Of(figure);
        switch (figure)
        {
            case IPoint:
                {
                    var x = (name + ".X", "the x of " + of);
                    var y = (name + ".Y", "the y of " + of);
                    var point = (name, "point " + of);
                    var order = context switch
                    {
                        InsertContext.PointX => new[] { x, y, point },
                        InsertContext.PointY => new[] { y, x, point },
                        _ => new[] { point, x, y }
                    };
                    foreach (var (text, description) in order)
                    {
                        Add(text, description);
                    }

                    break;
                }

            case Slider:
                Add(name, "slider " + figure.Name);
                break;
            case Number:
                Add(name, "number " + figure.Name);
                break;
            case Label label:
                if (label.IsNumber)
                {
                    Add(name + ".Value", "the value of the text");
                }

                break;
            case LengthMeasurement measurement:
                AddLength(measurement, Add);
                break;
            case AngleMeasurementBase:
            case AngleArc:
                Add(AngleText(figure), "angle " + AnglePoints(figure));
                break;
            case AreaMeasurement area:
                AddArea(area, Add);
                break;
            case Segment segment when segment.Dependencies.Count >= 2 && segment.Dependencies[0] is IPoint start && segment.Dependencies[1] is IPoint end:
                Add(TwoPoints(segment.Drawing, start, end), "the length of " + of);
                Add(name + ".Direction", "the direction of " + of);
                break;
            case Vector:
                Add(name + ".Length", "the length of " + of);
                Add(name + ".X", "the x of " + of);
                Add(name + ".Y", "the y of " + of);
                break;
            case CircleBase:
                AddQuantities(figure, name, of, Add);
                Add(name + ".Area", "the area of " + of);
                break;
            case EllipseArcBase:
                AddQuantities(figure, name, of, Add);
                if (figure is IShapeWithInterior)
                {
                    Add(name + ".Area", "the area of " + of);
                }

                break;
            case EllipseBase:
                Add(name + ".Area", "the area of " + of);
                Add(name + ".Circumference", "the circumference of " + of);
                break;
            case PolygonBase:
                Add(name + ".Area", "the area of " + of);
                Add(name + ".Perimeter", "the perimeter of " + of);
                break;
            case BezierPath path:
                Add(name + (path.Closed ? ".Perimeter" : ".Length"), (path.Closed ? "the perimeter of " : "the length of ") + of);
                if (path.HasInside)
                {
                    Add(name + ".Area", "the area of " + of);
                }

                break;
            case Polyline:
            case Bezier:
                Add(name + ".Length", "the length of " + of);
                break;
        }

        return options;
    }

    /// <summary>The lengths a figure has, as the Length tool offers them (<see cref="LengthMeasurement.QuantitiesOf"/>), the first being what a click measures</summary>
    static void AddQuantities(IFigure figure, string name, string of, System.Action<string, string> add)
    {
        foreach (var quantity in LengthMeasurement.QuantitiesOf(figure))
        {
            add(name + "." + PropertyOf(quantity), WordsOf(quantity) + " of " + of);
        }
    }

    /// <summary>A length measurement is said as what it measures: the figure's quantity, its own first, or the two points</summary>
    static void AddLength(LengthMeasurement measurement, System.Action<string, string> add)
    {
        if (measurement.Dependencies.Count == 2 && measurement.Dependencies[0] is IPoint first && measurement.Dependencies[1] is IPoint second)
        {
            add(TwoPoints(measurement.Drawing, first, second), "the distance " + first.Name + second.Name);
            return;
        }

        if (measurement.Dependencies.Count != 1)
        {
            return;
        }

        var measured = measurement.Dependencies[0];
        string name = Scanner.NameInExpression(measured.Name);
        string of = ConstructionText.Of(measured);
        var quantities = LengthMeasurement.QuantitiesOf(measured).ToList();
        if (quantities.Remove(measurement.Quantity))
        {
            quantities.Insert(0, measurement.Quantity);
        }

        foreach (var quantity in quantities)
        {
            add(name + "." + PropertyOf(quantity), WordsOf(quantity) + " of " + of);
        }

        if (measured is IShapeWithInterior)
        {
            add(name + ".Area", "the area of " + of);
        }
    }

    /// <summary>An area measurement is said as the area of what it measures: the shape's, or the function of the points</summary>
    static void AddArea(AreaMeasurement measurement, System.Action<string, string> add)
    {
        if (measurement.Dependencies.Count == 0)
        {
            return;
        }

        if (measurement.Dependencies[0] is IPoint)
        {
            var names = measurement.Dependencies.Select(point => Scanner.NameInExpression(point.Name));
            add("area(" + string.Join(", ", names) + ")", "the area of " + ConstructionText.Points(measurement.Dependencies.ToArray()));
            return;
        }

        var measured = measurement.Dependencies[0];
        add(Scanner.NameInExpression(measured.Name) + ".Area", "the area of " + ConstructionText.Of(measured));
    }

    /// <summary>The property an expression reads for a length of a figure (<see cref="LengthQuantity"/>)</summary>
    static string PropertyOf(LengthQuantity quantity)
    {
        return quantity == LengthQuantity.Distance ? "Length" : quantity.ToString();
    }

    static string WordsOf(LengthQuantity quantity)
    {
        return quantity switch
        {
            LengthQuantity.ArcLength => "the arc length",
            LengthQuantity.Radius => "the radius",
            LengthQuantity.Diameter => "the diameter",
            LengthQuantity.Circumference => "the circumference",
            LengthQuantity.Perimeter => "the perimeter",
            _ => "the length"
        };
    }

    /// <summary>
    /// The distance between two points: their names run together (AB), as a caption says it,
    /// where the language reads that as the two points; else dist(A, B) (a name that needs
    /// backticks, two names that run together into another reading)
    /// </summary>
    public static string TwoPoints(Drawing drawing, IPoint first, IPoint second)
    {
        string run = first.Name + second.Name;
        if (drawing != null && Scanner.IsName(first.Name) && Scanner.IsName(second.Name) && Scanner.IsName(run))
        {
            var result = Compiler.Instance.CompileExpression(drawing, run, isFigureAllowed: null);
            if (result.IsSuccess
                && result.Dependencies.Count == 2
                && result.Dependencies.Contains(first)
                && result.Dependencies.Contains(second))
            {
                return run;
            }
        }

        return "dist(" + Scanner.NameInExpression(first.Name) + ", " + Scanner.NameInExpression(second.Name) + ")";
    }

    /// <summary>
    /// An angle (a measurement, a mark) as its Sweep has it right now: the smaller angle is
    /// angle(A, B, C), the larger 360 less that, the oriented ones through oang; in radians
    /// when the measurement shows them. The vertex is the angle's first dependency.
    /// </summary>
    public static string AngleText(IFigure angle)
    {
        var dependencies = angle.Dependencies;
        if (dependencies.Count < 3)
        {
            return null;
        }

        string points = Scanner.NameInExpression(dependencies[1].Name)
            + ", " + Scanner.NameInExpression(dependencies[0].Name)
            + ", " + Scanner.NameInExpression(dependencies[2].Name);
        var sweep = angle is IHasSweep withSweep ? withSweep.Sweep : AngleSweep.Smaller;
        string text = sweep switch
        {
            AngleSweep.Larger => "360 - angle(" + points + ")",
            AngleSweep.Counterclockwise => "deg(oang(" + points + "))",
            AngleSweep.Clockwise => "360 - deg(oang(" + points + "))",
            _ => "angle(" + points + ")"
        };
        return angle is AngleMeasurementBase { Radians: true } ? "rad(" + text + ")" : text;
    }

    /// <summary>"ABC": the angle's points, the vertex in the middle</summary>
    static string AnglePoints(IFigure angle)
    {
        var dependencies = angle.Dependencies;
        return dependencies.Count < 3 ? "" : ConstructionText.Points(dependencies[1], dependencies[0], dependencies[2]);
    }
}
