using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;

namespace DynamicGeometry;

/// <summary>
/// The icon of a tool made by Define figure: a small picture of what the tool was defined on,
/// as it was on the paper when Create tool was pressed. The inputs are drawn as the ribbon's
/// icons draw the figures a tool starts from (yellow points, lines in ink), the figures made
/// as the ones a tool makes (constructed points, lines in the accent color). Taken once, in
/// the units of the icon, and kept in the macro (<c>&lt;Icon&gt;</c>), so that it doesn't
/// depend on the drawing. A figure with no simple picture (a label, a measurement, an angle's
/// mark) is left out; a tool with nothing to draw shows a dot and the number of its inputs.
/// </summary>
public static class MacroIcon
{
    // how many steps a circle, an ellipse, an arc or a Bézier curve is drawn in
    const int CurveSteps = 32;

    // room left around the picture, in the icon's units
    const double Margin = 0.1;

    class Stroke
    {
        public List<Point> Points = new List<Point>();
        public bool IsMade;
        public bool IsClosed;
        public bool IsFilled;

        // a line or a ray, or a function graph: as far as the window, so it says nothing about
        // where the picture is, and is cut to the picture's box
        public bool IsUnbounded;
    }

    class Dot
    {
        public Point Coordinates;
        public bool IsMade;
    }

    public static void Write(XmlWriter writer, IList<IFigure> inputs, IList<IFigure> results)
    {
        var strokes = new List<Stroke>();
        var dots = new List<Dot>();
        foreach (var input in inputs)
        {
            Collect(input, isMade: false, strokes, dots);
        }

        foreach (var result in results)
        {
            Collect(result, isMade: true, strokes, dots);
        }

        var drawing = inputs.Concat(results).Select(f => f.Drawing).FirstOrDefault(d => d != null);
        var box = FindBox(strokes, dots, drawing);
        if (box == null)
        {
            return;
        }

        writer.WriteStartElement("Icon");
        foreach (var stroke in strokes)
        {
            foreach (var piece in Clip(stroke.Points.Select(p => ToIcon(p, box.Value)).ToList(), stroke.IsClosed))
            {
                writer.WriteStartElement("Stroke");
                writer.WriteAttributeString("Points", string.Join(" ", piece.Select(p => Format(p.X) + "," + Format(p.Y))));
                WriteFlag(writer, "Made", stroke.IsMade);
                WriteFlag(writer, "Closed", stroke.IsClosed && piece.Count == stroke.Points.Count);
                WriteFlag(writer, "Filled", stroke.IsFilled && piece.Count == stroke.Points.Count);
                writer.WriteEndElement();
            }
        }

        foreach (var dot in dots)
        {
            var place = ToIcon(dot.Coordinates, box.Value);
            if (place.X < 0 || place.X > 1 || place.Y < 0 || place.Y > 1)
            {
                continue;
            }

            writer.WriteStartElement("Dot");
            writer.WriteAttributeString("X", Format(place.X));
            writer.WriteAttributeString("Y", Format(place.Y));
            WriteFlag(writer, "Made", dot.IsMade);
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
    }

    /// <summary>The icon a macro's <c>&lt;Icon&gt;</c> describes; a dot and the number of inputs when it has none</summary>
    public static FrameworkElement Build(XElement macro)
    {
        var icon = macro?.Element("Icon");
        var builder = IconBuilder.BuildIcon();
        if (icon == null || !icon.HasElements)
        {
            int inputCount = macro?.Element("Inputs")?.Elements().Count() ?? 0;
            return builder
                .Point(0.3, 0.5)
                .Text(nameof(AppTheme.Ink), 0.5, 0.2, inputCount.ToString(CultureInfo.InvariantCulture), fontSize: 14)
                .Canvas;
        }

        // what the tool starts from under what it makes, the points on top
        var strokes = icon.Elements("Stroke").OrderBy(s => s.ReadBool("Made", defaultValue: false)).ToList();
        foreach (var stroke in strokes)
        {
            var points = ParsePoints(stroke.ReadString("Points"));
            if (points.Count < 2)
            {
                continue;
            }

            bool isMade = stroke.ReadBool("Made", defaultValue: false);
            string color = isMade ? nameof(AppTheme.LineAccent) : nameof(AppTheme.Ink);
            double thickness = isMade ? IconBuilder.AccentThickness : 1;
            if (stroke.ReadBool("Filled", defaultValue: false))
            {
                builder.Polygon(nameof(AppTheme.ShapeIconFill), color, points.ToArray());
            }
            else
            {
                builder.Polyline(thickness, color, points, isClosed: stroke.ReadBool("Closed", defaultValue: false));
            }
        }

        foreach (var dot in icon.Elements("Dot"))
        {
            double x = dot.ReadDouble("X");
            double y = dot.ReadDouble("Y");
            if (dot.ReadBool("Made", defaultValue: false))
            {
                builder.DependentPoint(x, y);
            }
            else
            {
                builder.Point(x, y);
            }
        }

        return builder.Canvas;
    }

    static void Collect(IFigure figure, bool isMade, List<Stroke> strokes, List<Dot> dots)
    {
        if (figure == null || !figure.Visible || !figure.Exists || figure is LabelBase || figure is AngleArc)
        {
            return;
        }

        switch (figure)
        {
            case IPoint point:
                dots.Add(new Dot { Coordinates = point.Coordinates, IsMade = isMade });
                return;
            case ILine line:
                var ends = line is LineBase lineBase ? lineBase.OnScreenCoordinates : line.Coordinates;
                strokes.Add(new Stroke
                {
                    Points = { ends.P1, ends.P2 },
                    IsMade = isMade,
                    IsUnbounded = !(line is Segment || line is Vector)
                });
                return;
            case IPolygonalChain chain when chain.VertexCoordinates != null:
                bool hasInterior = figure is IShapeWithInterior;
                strokes.Add(new Stroke
                {
                    Points = chain.VertexCoordinates.ToList(),
                    IsMade = isMade,
                    IsClosed = hasInterior,
                    IsFilled = hasInterior
                });
                return;
            case Curve curve:
                var samples = new List<Point>();
                curve.GetPoints(samples);
                var stretch = new Stroke { IsMade = isMade, IsUnbounded = curve is FunctionGraph };
                foreach (var sample in samples)
                {
                    if (sample.Exists())
                    {
                        stretch.Points.Add(sample);
                        continue;
                    }

                    // a gap: what comes after it is another stretch
                    if (stretch.Points.Count > 1)
                    {
                        strokes.Add(stretch);
                    }

                    stretch = new Stroke { IsMade = isMade, IsUnbounded = stretch.IsUnbounded };
                }

                if (stretch.Points.Count > 1)
                {
                    strokes.Add(stretch);
                }

                return;
            case ILinearFigure linear when figure is IEllipse || figure is Bezier:
                var domain = linear.GetParameterDomain();
                var stroke = new Stroke { IsMade = isMade, IsClosed = figure is IEllipse && !(figure is IArc) };
                for (int i = 0; i <= CurveSteps; i++)
                {
                    var sample = linear.GetPointFromParameter(domain.Item1 + (domain.Item2 - domain.Item1) * i / CurveSteps);
                    if (sample.Exists())
                    {
                        stroke.Points.Add(sample);
                    }
                }

                if (stroke.Points.Count > 1)
                {
                    strokes.Add(stroke);
                }

                return;
            case CompositeFigure composite:
                // a slider: its anchor, knob and track
                foreach (var child in composite.Children)
                {
                    Collect(child, isMade, strokes, dots);
                }

                return;
        }
    }

    /// <summary>
    /// The square of the plane the icon shows: around the points and the figures that end
    /// somewhere, else the window. Null when there is nothing to draw.
    /// </summary>
    static Rect? FindBox(List<Stroke> strokes, List<Dot> dots, Drawing drawing)
    {
        var places = dots.Select(d => d.Coordinates)
            .Concat(strokes.Where(s => !s.IsUnbounded).SelectMany(s => s.Points))
            .ToList();
        if (places.Count == 0)
        {
            if (strokes.Count == 0 || drawing == null)
            {
                return null;
            }

            var view = drawing.CoordinateSystem;
            places.Add(new Point(view.MinimalVisibleX, view.MinimalVisibleY));
            places.Add(new Point(view.MaximalVisibleX, view.MaximalVisibleY));
        }

        double left = places.Min(p => p.X);
        double right = places.Max(p => p.X);
        double bottom = places.Min(p => p.Y);
        double top = places.Max(p => p.Y);

        // a square, so that a circle stays round; a single point gets a square of its own
        double size = System.Math.Max(right - left, top - bottom);
        if (size <= 0)
        {
            size = 1;
        }

        double middleX = (left + right) / 2;
        double middleY = (bottom + top) / 2;
        return new Rect(middleX - size / 2, middleY - size / 2, size, size);
    }

    /// <summary>A point of the plane in the icon's units: 0 to 1 across the box, y going down</summary>
    static Point ToIcon(Point point, Rect box)
    {
        double scale = 1 - 2 * Margin;
        return new Point(
            Margin + (point.X - box.X) / box.Width * scale,
            Margin + (box.Bottom - point.Y) / box.Height * scale);
    }

    /// <summary>
    /// The stretches of a polyline inside the icon: a line through the picture runs to its
    /// edges, a graph's points far off are left out
    /// </summary>
    static List<List<Point>> Clip(List<Point> points, bool isClosed)
    {
        var path = isClosed && points.Count > 2 ? points.Append(points[0]).ToList() : points;
        var pieces = new List<List<Point>>();
        List<Point> current = null;
        for (int i = 0; i + 1 < path.Count; i++)
        {
            if (!ClipSegment(path[i], path[i + 1], out var start, out var end))
            {
                current = null;
                continue;
            }

            if (current == null || current[current.Count - 1] != start)
            {
                current = new List<Point> { start };
                pieces.Add(current);
            }

            current.Add(end);
        }

        // whole and closed: the same points as given, the closing side drawn by the shape
        if (isClosed && pieces.Count == 1 && pieces[0].Count == path.Count && pieces[0].SequenceEqual(path))
        {
            return new List<List<Point>> { points };
        }

        return pieces;
    }

    /// <summary>The part of the segment from a to b inside the unit square (Liang-Barsky)</summary>
    static bool ClipSegment(Point a, Point b, out Point start, out Point end)
    {
        double t0 = 0;
        double t1 = 1;
        double dx = b.X - a.X;
        double dy = b.Y - a.Y;
        var edges = new[] { (-dx, a.X), (dx, 1 - a.X), (-dy, a.Y), (dy, 1 - a.Y) };
        foreach (var (p, q) in edges)
        {
            if (p == 0)
            {
                if (q < 0)
                {
                    start = end = default;
                    return false;
                }

                continue;
            }

            double t = q / p;
            if (p < 0)
            {
                t0 = System.Math.Max(t0, t);
            }
            else
            {
                t1 = System.Math.Min(t1, t);
            }
        }

        start = t0 == 0 ? a : new Point(a.X + t0 * dx, a.Y + t0 * dy);
        end = t1 == 1 ? b : new Point(a.X + t1 * dx, a.Y + t1 * dy);
        return t0 <= t1;
    }

    static List<Point> ParsePoints(string text)
    {
        var points = new List<Point>();
        if (text == null)
        {
            return points;
        }

        foreach (var pair in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = pair.Split(',');
            if (parts.Length == 2
                && double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var x)
                && double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var y))
            {
                points.Add(new Point(x, y));
            }
        }

        return points;
    }

    static string Format(double value)
    {
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }

    static void WriteFlag(XmlWriter writer, string name, bool value)
    {
        if (value)
        {
            writer.WriteAttributeString(name, "true");
        }
    }
}
