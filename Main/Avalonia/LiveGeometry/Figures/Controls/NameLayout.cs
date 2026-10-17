using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// Names that keep clear (2026-10-17): a point's name or a figure's name is drawn where its
/// offset puts it - where it was left, or where it went first - unless something covers that
/// place: another name, a measurement, an angle's number or another point. Then it is drawn
/// at the nearest free place around its anchor, at the same distance or a little further,
/// and not across the lines through the anchor (a point's own lines; a segment's name, its
/// segment), until its own place is free again. Labels keep their size in pixels whatever
/// the zoom, so names that sat apart on a big screen met on a phone (P and the r of the
/// segment beside it). Only where the name is drawn changes: its offset, and what the file
/// says, stay where the user left them. A name stays where it went while that is free, so
/// it doesn't flicker between two places during a drag, and a name being dragged is drawn
/// under the cursor whatever it covers. The player has the twin (nameLayout.js).
/// </summary>
public static class NameLayout
{
    /// <summary>In square pixels: an overlap this small is a touch, not a cover</summary>
    public const double OverlapAllowance = 2;

    /// <summary>In degrees: the steps around the anchor a name tries, the nearest first</summary>
    public const double StepDegrees = 15;

    /// <summary>The steps each way: all the way round</summary>
    const int Steps = 12;

    /// <summary>In pixels: kept between a name and a point, and between a name and its own point</summary>
    public const double PointGap = 2;

    /// <summary>In pixels: kept between a name that has moved and the other names, so that A and B side by side don't read as AB (at 2 they did)</summary>
    public const double Spacing = 6;

    /// <summary>In pixels: how far into a name at its own place another name may reach, both ways, before it moves</summary>
    public const double HomeDepth = 2;

    /// <summary>In pixels: the size of a point's marker that hasn't been measured</summary>
    const double DefaultPointSize = 8;

    /// <summary>In pixels: how near the anchor a line must run to be one through it</summary>
    const double OnAnchor = 1.5;

    // in steps around the anchor: a square pixel covered, a pixel of a line through the name
    const double OverlapCost = 0.5;
    const double CrossingCost = 0.5;

    // the names of a drawing drawn away from their place, which go back once it is free
    static readonly ConditionalWeakTable<Drawing, HashSet<LabelWithOffset>> nudged = new();

    static bool rechecking;

    static HashSet<LabelWithOffset> Nudged(Drawing drawing)
    {
        return nudged.GetValue(drawing, _ => new HashSet<LabelWithOffset>());
    }

    /// <summary>
    /// Where the name is drawn, its top-left corner in pixels, given its place there: the
    /// place while nothing covers it, else the nearest free place around the anchor, which
    /// <see cref="LabelWithOffset.Nudge"/> remembers
    /// </summary>
    public static Point Place(LabelWithOffset label, Point place)
    {
        return Choose(label, place, isFirst: false);
    }

    /// <summary>
    /// Where a new name goes, its top-left corner in pixels, given its default place: that,
    /// unless something covers it or a line through the anchor runs across it (a point's
    /// name above it on a vertical line), else the nearest free place around the anchor.
    /// The result becomes the name's offset; nothing is remembered.
    /// </summary>
    public static Point FirstPlace(LabelWithOffset label, Point place)
    {
        return Choose(label, place, isFirst: true);
    }

    static Point Choose(LabelWithOffset label, Point place, bool isFirst)
    {
        var drawing = label.Drawing;
        var size = label.ShownSize();
        if (!(size.Width > 0 && size.Height > 0) || !place.Exists())
        {
            Forget(label, drawing);
            return place;
        }

        var coordinateSystem = drawing.CoordinateSystem;
        var anchor = coordinateSystem.ToPhysical(label.Anchor);
        var home = new Rect(place, size);
        if (!anchor.Exists())
        {
            Forget(label, drawing);
            return place;
        }

        // what is near: the other names and the points within reach of every place tried
        double distance = home.Center.Distance(anchor);
        double reach = distance + 2 * System.Math.Max(size.Width, size.Height) + DefaultPointSize;
        var ownPoint = label.OwnPoint;
        Rect? ownMarker = null;
        var points = new List<Rect>();
        var names = new List<Rect>();
        foreach (var figure in drawing.Figures)
        {
            if (figure is LabelWithOffset other && other != label)
            {
                if (other.ShownBox() is Rect name && DistanceToBox(anchor, name) <= reach)
                {
                    names.Add(name);
                }
            }
            else if (figure is PointBase point && MarkerBox(point, coordinateSystem) is Rect marker && DistanceToBox(anchor, marker) <= reach)
            {
                if (figure == ownPoint)
                {
                    ownMarker = marker;
                }
                else
                {
                    points.Add(marker);
                }
            }
        }

        // a place it goes to keeps Spacing from the other names
        var obstacles = points.Concat(names.Select(name => name.Inflate(Spacing / 2))).ToList();

        // a place the user left is theirs, lines across it or not; a new one keeps off them
        var lines = new List<(Point From, Point To)>();
        bool linesFound = false;
        double Crossing(Rect box)
        {
            if (!linesFound)
            {
                linesFound = true;
                foreach (var figure in label.FiguresThroughAnchor().Distinct())
                {
                    AddLine(figure, coordinateSystem, anchor, reach, lines);
                }
            }

            return lines.Sum(line => ClippedLength(line.From, line.To, box.Deflate(1)));
        }

        // A name's own place is covered by a point over it, or by a name that overlaps it by
        // more than HomeDepth both ways: the boxes have room around the letters, and a
        // clock's numbers whose boxes touch on a phone are left where they are
        double homeOverlap = Overlap(home, points) + names.Sum(name => Deep(home.Intersect(name)));
        double homeCrossing = isFirst && homeOverlap <= OverlapAllowance ? Crossing(home) : 0;
        if (homeOverlap <= OverlapAllowance && homeCrossing == 0)
        {
            return Forgotten(label, drawing, place);
        }

        double Cost(Rect box, int steps, out bool isFree)
        {
            double overlap = Overlap(box, obstacles) + (ownMarker != null ? Area(box.Intersect(ownMarker.Value)) : 0);
            double crossing = Crossing(box);
            isFree = overlap <= OverlapAllowance && crossing == 0;
            return steps + overlap * OverlapCost + crossing * CrossingCost;
        }

        // where it went last time, while that is still free
        if (!isFirst && label.Nudge is Point nudge)
        {
            var kept = Centered(anchor + nudge, size);
            Cost(kept, 0, out bool keptIsFree);
            if (keptIsFree)
            {
                return kept.TopLeft;
            }
        }

        // around the anchor, the nearest first: at the place's distance, then a ring further out
        var direction = home.Center - anchor;
        if (direction.Length() < 1)
        {
            direction = new Point(0, -1);
        }

        direction *= 1 / direction.Length();
        double ringStep = System.Math.Max(size.Width, size.Height) / 2 + 2 * PointGap;
        Rect best = home;
        double bestCost = homeOverlap * OverlapCost + homeCrossing * CrossingCost;
        for (int ring = 0; ring < 2; ring++)
        {
            for (int step = 1; step <= Steps; step++)
            {
                foreach (int sign in new[] { 1, -1 })
                {
                    var turned = Rotate(direction, sign * step * StepDegrees);
                    var box = PlaceOut(anchor, turned, distance + ring * ringStep, size, ownMarker);
                    double cost = Cost(box, step + ring * Steps, out bool isFree);
                    if (isFree)
                    {
                        return isFirst ? box.TopLeft : Remember(label, drawing, anchor, box);
                    }

                    if (cost < bestCost)
                    {
                        bestCost = cost;
                        best = box;
                    }
                }
            }
        }

        // nowhere free: the place that covers least, which may be its own
        if (isFirst)
        {
            return best.TopLeft;
        }

        return best == home ? Forgotten(label, drawing, place) : Remember(label, drawing, anchor, best);
    }

    static Point Remember(LabelWithOffset label, Drawing drawing, Point anchor, Rect box)
    {
        label.Nudge = box.Center - anchor;
        Nudged(drawing).Add(label);
        return box.TopLeft;
    }

    static Point Forgotten(LabelWithOffset label, Drawing drawing, Point place)
    {
        Forget(label, drawing);
        return place;
    }

    /// <summary>The name is at its own place again, or has left the drawing</summary>
    public static void Forget(LabelWithOffset label, Drawing drawing)
    {
        label.Nudge = null;
        if (drawing != null && nudged.TryGetValue(drawing, out var names))
        {
            names.Remove(label);
        }
    }

    /// <summary>
    /// A label was drawn somewhere else: the names drawn away from their place look again,
    /// since their place may be free now (asked by every label that moves, not again from
    /// inside: one look per move)
    /// </summary>
    public static void Recheck(LabelWithOffset moved)
    {
        var drawing = moved.Drawing;
        if (rechecking || drawing == null || !nudged.TryGetValue(drawing, out var names) || names.Count == 0)
        {
            return;
        }

        rechecking = true;
        try
        {
            foreach (var name in names.ToArray())
            {
                if (name != moved && name.Drawing == drawing)
                {
                    name.UpdateVisual();
                }
            }
        }
        finally
        {
            rechecking = false;
        }
    }

    /// <summary>A shown point's marker on the canvas, in pixels, with <see cref="PointGap"/> around it</summary>
    static Rect? MarkerBox(PointBase point, CoordinateSystem coordinateSystem)
    {
        if (!point.Visible || !point.Exists || point.Shape == null)
        {
            return null;
        }

        var center = coordinateSystem.ToPhysical(point.Coordinates);
        if (!center.Exists())
        {
            return null;
        }

        var measured = point.Shape.Bounds.Size;
        double width = measured.Width > 0 ? measured.Width : DefaultPointSize;
        double height = measured.Height > 0 ? measured.Height : DefaultPointSize;
        return new Rect(
            center.X - width / 2 - PointGap,
            center.Y - height / 2 - PointGap,
            width + 2 * PointGap,
            height + 2 * PointGap);
    }

    /// <summary>
    /// A line or circle that runs through the anchor, as pieces in pixels near it: a
    /// segment, a ray or a line as far as the reach, a circle as chords. Nothing for a
    /// figure that doesn't run through the anchor (a circle around a point).
    /// </summary>
    static void AddLine(IFigure figure, CoordinateSystem coordinateSystem, Point anchor, double reach, List<(Point From, Point To)> lines)
    {
        if (!figure.Visible || !figure.Exists)
        {
            return;
        }

        if (figure is ILine line)
        {
            var a = coordinateSystem.ToPhysical(line.Coordinates.P1);
            var b = coordinateSystem.ToPhysical(line.Coordinates.P2);
            var along = b - a;
            double length = along.Length();
            if (!(length > 0) || !a.Exists() || !b.Exists())
            {
                return;
            }

            var unit = along * (1 / length);
            double far = (anchor - a).Length() + reach;
            var from = a;
            var to = b;
            if (figure is Ray)
            {
                to = a + unit * far;
            }
            else if (figure is not Segment && figure is not Vector)
            {
                from = a - unit * far;
                to = a + unit * far;
            }

            if (DistanceToSegment(anchor, from, to) <= OnAnchor)
            {
                lines.Add((from, to));
            }

            return;
        }

        if (figure is ICircle circle)
        {
            var center = coordinateSystem.ToPhysical(circle.Center);
            double radius = circle.Radius * coordinateSystem.UnitLength;
            if (!(radius > 0) || !center.Exists() || System.Math.Abs((anchor - center).Length() - radius) > OnAnchor)
            {
                return;
            }

            // the chords near the anchor, two degrees apart
            const int chords = 180;
            var previous = center + new Point(radius, 0);
            for (int i = 1; i <= chords; i++)
            {
                double angle = 2 * System.Math.PI * i / chords;
                var next = center + new Point(radius * System.Math.Cos(angle), radius * System.Math.Sin(angle));
                if (DistanceToSegment(anchor, previous, next) <= reach)
                {
                    lines.Add((previous, next));
                }

                previous = next;
            }
        }
    }

    /// <summary>A box of the size centered on the direction from the anchor at the distance, pushed out until it clears the anchor's own point</summary>
    static Rect PlaceOut(Point anchor, Point direction, double distance, Size size, Rect? ownMarker)
    {
        var box = Centered(anchor + direction * distance, size);
        if (ownMarker == null)
        {
            return box;
        }

        double limit = distance + System.Math.Max(size.Width, size.Height);
        while (box.Intersects(ownMarker.Value) && distance < limit)
        {
            distance += 1;
            box = Centered(anchor + direction * distance, size);
        }

        return box;
    }

    static Rect Centered(Point center, Size size)
    {
        return new Rect(center.X - size.Width / 2, center.Y - size.Height / 2, size.Width, size.Height);
    }

    static Point Rotate(Point vector, double degrees)
    {
        double radians = degrees * System.Math.PI / 180;
        double cos = System.Math.Cos(radians);
        double sin = System.Math.Sin(radians);
        return new Point(vector.X * cos - vector.Y * sin, vector.X * sin + vector.Y * cos);
    }

    static double Overlap(Rect box, List<Rect> obstacles)
    {
        double sum = 0;
        foreach (var obstacle in obstacles)
        {
            sum += Area(box.Intersect(obstacle));
        }

        return sum;
    }

    static double Area(Rect rect) => rect.Width > 0 && rect.Height > 0 ? rect.Width * rect.Height : 0;

    /// <summary>The area of an overlap deeper than <see cref="HomeDepth"/> both ways; none for a shallower one</summary>
    static double Deep(Rect overlap) => overlap.Width > HomeDepth && overlap.Height > HomeDepth ? overlap.Width * overlap.Height : 0;

    static double DistanceToBox(Point point, Rect box)
    {
        double dx = System.Math.Max(System.Math.Max(box.Left - point.X, point.X - box.Right), 0);
        double dy = System.Math.Max(System.Math.Max(box.Top - point.Y, point.Y - box.Bottom), 0);
        return System.Math.Sqrt(dx * dx + dy * dy);
    }

    static double DistanceToSegment(Point point, Point from, Point to)
    {
        var along = to - from;
        double squared = along.X * along.X + along.Y * along.Y;
        double t = squared > 0 ? ((point.X - from.X) * along.X + (point.Y - from.Y) * along.Y) / squared : 0;
        t = System.Math.Clamp(t, 0, 1);
        return (point - (from + along * t)).Length();
    }

    /// <summary>How much of the piece lies in the box, in pixels (Liang-Barsky)</summary>
    public static double ClippedLength(Point from, Point to, Rect box)
    {
        double dx = to.X - from.X;
        double dy = to.Y - from.Y;
        double low = 0;
        double high = 1;
        if (!Clip(-dx, from.X - box.Left, ref low, ref high)
            || !Clip(dx, box.Right - from.X, ref low, ref high)
            || !Clip(-dy, from.Y - box.Top, ref low, ref high)
            || !Clip(dy, box.Bottom - from.Y, ref low, ref high))
        {
            return 0;
        }

        return (high - low) * System.Math.Sqrt(dx * dx + dy * dy);
    }

    static bool Clip(double p, double q, ref double low, ref double high)
    {
        if (p == 0)
        {
            return q >= 0;
        }

        double r = q / p;
        if (p < 0)
        {
            if (r > high)
            {
                return false;
            }

            low = System.Math.Max(low, r);
        }
        else
        {
            if (r < low)
            {
                return false;
            }

            high = System.Math.Min(high, r);
        }

        return true;
    }
}
