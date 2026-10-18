using System;
using System.Collections.Generic;
using System.Linq;
using GuiLabs.Undo;

namespace LiveGeometry;

/// <summary>What a <see cref="FigureQuantity"/> is counted in, for how it is written</summary>
public enum QuantityUnit
{
    /// <summary>A length or an area: the number as it is</summary>
    Number,

    /// <summary>Degrees, with a degree sign</summary>
    Degrees
}

/// <summary>
/// One thing that can be measured of a figure, for its page of the property grid: its
/// radius, its area, a line's slope. Read-only, written as the grid writes numbers, and
/// with a way to put it on the drawing as a measurement figure (the Length, Area and Angle
/// tools' figures) where there is one for it and the drawing hasn't got it yet.
/// </summary>
public class FigureQuantity
{
    public FigureQuantity(string caption, Func<double> value, QuantityUnit unit = QuantityUnit.Number)
    {
        Caption = caption;
        Value = value;
        Unit = unit;
    }

    public string Caption { get; }

    public Func<double> Value { get; }

    public QuantityUnit Unit { get; }

    /// <summary>
    /// The row of the figure's page that says this already, editable (a segment's Length, a
    /// circle's radius); null for none. The page then shows no second row for it, and the
    /// button that measures it goes on that row.
    /// </summary>
    public string RowName { get; init; }

    /// <summary>Whether the drawing has a measurement of this already; null where there is no measurement figure for it (a slope)</summary>
    public Func<bool> IsMeasured { get; init; }

    /// <summary>Puts a measurement of it on the drawing, one undo step</summary>
    public Action Measure { get; init; }

    /// <summary>What the "+" says it does</summary>
    public string MeasureHint => "Show the " + Caption.ToLowerInvariant() + " on the drawing";

    /// <summary>Whether a measurement of it can be added now</summary>
    public bool CanMeasure => Measure != null && IsMeasured != null && !IsMeasured() && Value().IsValidValue();

    /// <summary>The value as the grid writes numbers ("12.57", "45°"), "undefined" without one</summary>
    public string Text
    {
        get
        {
            double value = Value();
            if (!value.IsValidValue())
            {
                return LabelBase.UndefinedText;
            }

            string text = Math.Round(value, Settings.DisplayDecimals).ToString();
            return Unit == QuantityUnit.Degrees ? text + "°" : text;
        }
    }
}

/// <summary>
/// Whatever can be measured of a figure, listed for its page: the lengths the Length tool
/// can take of it (a circle's radius, diameter and circumference; a polygon's perimeter;
/// an arc's length), its area where the Area tool takes it, the central angle of an arc, a
/// sector or a segment, and a line's slope and angle. The first three can be put on the
/// drawing as measurements; a slope and a line's angle have no figure that says them.
/// </summary>
public static class FigureQuantities
{
    public static IReadOnlyList<FigureQuantity> Of(IFigure figure)
    {
        var result = new List<FigureQuantity>();
        if (figure?.Drawing == null)
        {
            return result;
        }

        AddLengths(figure, result);
        AddArea(figure, result);
        AddCentralAngle(figure, result);
        AddSlope(figure, result);
        return result;
    }

    static void AddLengths(IFigure figure, List<FigureQuantity> result)
    {
        // what the Length tool takes (LengthMeasurement.IsMeasurable), also while it has no
        // length right now: not a number, a label or a measurement, lengths with nothing to
        // measure, nor an angle's mark, an arc sized in pixels
        bool measurable = figure is Segment or Vector or Polyline or Bezier or BezierPath or EllipseArcBase or EllipseBase or PolygonBase
            && figure is not AngleArc;
        if (!measurable)
        {
            return;
        }

        foreach (var quantity in LengthMeasurement.QuantitiesOf(figure))
        {
            var kind = quantity;
            result.Add(new FigureQuantity(LengthMeasurement.KindOf(kind), () => LengthMeasurement.ValueOf(figure, kind))
            {
                RowName = LengthRowName(figure, kind),
                IsMeasured = () => HasLengthMeasurement(figure, kind),
                Measure = () => Actions.Add(figure.Drawing, Factory.CreateLengthMeasurement(figure.Drawing, new[] { figure }, kind))
            });
        }
    }

    /// <summary>The row that says the length already: a segment's or vector's Length, a circle's radius (its Length, captioned Radius)</summary>
    static string LengthRowName(IFigure figure, LengthQuantity quantity)
    {
        bool editable = quantity == LengthQuantity.Length && (figure is Segment || figure is Vector)
            || quantity == LengthQuantity.Radius && figure is CircleBase;
        return editable ? nameof(Segment.Length) : null;
    }

    /// <summary>
    /// A length measurement of the figure saying this quantity, or for a segment's length
    /// the distance between its two points, which is what the Length tool makes of two clicks
    /// </summary>
    static bool HasLengthMeasurement(IFigure figure, LengthQuantity quantity)
    {
        var ends = quantity == LengthQuantity.Length && figure is Segment ? figure.Dependencies : null;
        return figure.Drawing.Figures.OfType<LengthMeasurement>().Any(measurement =>
            measurement.Dependencies.Count == 1 && measurement.Dependencies[0] == figure && measurement.Quantity == quantity
            || ends != null && measurement.Dependencies.Count == 2 && ends.Count == 2 && SameTwo(measurement.Dependencies, ends));
    }

    static bool SameTwo(IList<IFigure> one, IList<IFigure> other)
    {
        return one[0] == other[0] && one[1] == other[1] || one[0] == other[1] && one[1] == other[0];
    }

    static void AddArea(IFigure figure, List<FigureQuantity> result)
    {
        if (!AreaMeasurement.IsMeasurable(figure))
        {
            return;
        }

        result.Add(new FigureQuantity("Area", () => AreaOf(figure))
        {
            IsMeasured = () => figure.Drawing.Figures.OfType<AreaMeasurement>().Any(measurement =>
                measurement.Dependencies.Count == 1 && measurement.Dependencies[0] == figure
                || figure is PolygonBase && measurement.Dependencies.SequenceEqual(figure.Dependencies)),
            Measure = () => Actions.Add(figure.Drawing, Factory.CreateAreaMeasurement(figure.Drawing, new[] { figure }))
        });
    }

    static double AreaOf(IFigure figure)
    {
        switch (figure)
        {
            case IShapeWithInterior shape:
                return shape.Area;
            case BezierPath path:
                return path.Area;
            default:
                return double.NaN;
        }
    }

    /// <summary>The angle at the center of an arc, a sector or a segment, the way round it goes (<see cref="AnglePoints"/>), as the Angle tool measures it</summary>
    static void AddCentralAngle(IFigure figure, List<FigureQuantity> result)
    {
        if (!AnglePoints.TakesArc(figure))
        {
            return;
        }

        result.Add(new FigureQuantity("Central angle", () => CentralAngle(figure), QuantityUnit.Degrees)
        {
            IsMeasured = () =>
            {
                var angle = AnglePoints.From(figure);
                return angle != null && figure.Drawing.Figures.OfType<AngleMeasurement>().Any(measurement =>
                    measurement.Dependencies.Count == 3
                    && measurement.Dependencies[0] == angle.Vertex
                    && SameTwo(new[] { measurement.Dependencies[1], measurement.Dependencies[2] }, new[] { angle.First, angle.Second }));
            },
            Measure = () => MeasureAngle(figure)
        });
    }

    static double CentralAngle(IFigure figure)
    {
        var angle = AnglePoints.From(figure);
        if (angle == null)
        {
            return double.NaN;
        }

        var points = angle.Points.ToPoints().ToArray();
        return angle.Sweep.Measure(Math.OAngle(points[1], points[0], points[2])).ToDegrees();
    }

    /// <summary>The mark and the number, as the Angle tool makes them of a click on an arc, one undo step</summary>
    static void MeasureAngle(IFigure figure)
    {
        var angle = AnglePoints.From(figure);
        if (angle == null)
        {
            return;
        }

        var drawing = figure.Drawing;
        var arc = Factory.CreateAngleArc(drawing, angle.Points);
        var measurement = Factory.CreateAngleMeasurement(drawing, angle.Points);
        arc.Sweep = angle.Sweep;
        measurement.Sweep = angle.Sweep;
        using (Transaction.Create(drawing.ActionManager, false))
        {
            Actions.Add(drawing, arc);
            Actions.Add(drawing, measurement);
        }
    }

    /// <summary>
    /// A straight figure's slope, and its angle where no row of its page says it (a
    /// segment's, a ray's and a line's Direction does; not a parallel's, a perpendicular's
    /// or a bisector's): from the x-axis counterclockwise, under 180° for a whole line
    /// </summary>
    static void AddSlope(IFigure figure, List<FigureQuantity> result)
    {
        if (!(figure is ILine line) || figure is AxisLine)
        {
            return;
        }

        result.Add(new FigureQuantity("Slope", () => Slope(line)));
        result.Add(new FigureQuantity("Angle", () => Angle(line), QuantityUnit.Degrees) { RowName = "Direction" });
    }

    static double Slope(ILine line)
    {
        var (from, to) = (line.Coordinates.P1, line.Coordinates.P2);
        double dx = to.X - from.X;
        double dy = to.Y - from.Y;

        // vertical (or no line at all): no slope, rather than a huge one from rounding
        if (System.Math.Abs(dx) <= 1e-12 * (System.Math.Abs(dy) + 1))
        {
            return double.NaN;
        }

        return dy / dx;
    }

    static double Angle(ILine line)
    {
        var (from, to) = (line.Coordinates.P1, line.Coordinates.P2);
        if (from == to)
        {
            return double.NaN;
        }

        double degrees = System.Math.Atan2(to.Y - from.Y, to.X - from.X).ToDegrees();
        double turn = line is Ray || line is Segment || line is Vector ? 360 : 180;
        degrees %= turn;
        if (degrees < 0)
        {
            degrees += turn;
        }

        // 179.9999999 is a line at 0
        return Math.Round(degrees, 9) == turn ? 0 : degrees;
    }
}
