using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// Where the numbers of the angles at a vertex go while nobody has placed them by hand
/// (<see cref="AngleMeasurement.PlacesItself"/>): each inside its own angle, past its mark,
/// and clear of the other numbers at the vertex as far as there is room. Worked out in
/// pixels whenever a number is drawn, for all the numbers at its vertex at once: a new
/// angle measured beside others moves them apart, and one placed by hand is an obstacle the
/// others keep clear of.
///
/// A place is a turn off the bisector (a fraction of the half angle, up to three quarters of
/// the way to a side) and a push out, past where the box clears the mark. A number tries a
/// few turns, each with pushes step by step, and takes the place that costs the least: the
/// area it overlaps other numbers by counts most, then how much of the sides of the angles
/// at the vertex runs through it, its middle lying outside its own angle, the turn and the
/// push. All but the middle outside go smoothly with the figures.
///
/// A number remembers its place (<see cref="Placement"/>, relative to its angle, so it turns
/// with it) and starts from there at the next redraw: it slides by small steps to a cheaper
/// place nearby, and goes to one further away only when that is cheaper by
/// <see cref="Hysteresis"/> and by <see cref="Proximity"/> per pixel of the way. Without
/// that, two places far apart traded being the cheapest as a point was dragged, and the
/// number jumped between them at every move. Two passes, every number starting from its
/// remembered place: the second puts every number again with all the others in theirs.
/// The player has the twin (angleLabelLayout.js).
/// </summary>
public static class AngleLabelLayout
{
    /// <summary>In pixels: between a mark and its number</summary>
    public const double Gap = 3;

    /// <summary>In pixels: kept between two numbers</summary>
    public const double Spacing = 2;

    /// <summary>In pixels: how far a number is pushed out past where it clears its mark</summary>
    public const double MaxPush = 90;

    /// <summary>In pixels: the steps of that push</summary>
    public const double PushStep = 3;

    /// <summary>In pixels: what a number keeps clear of the vertex when its angle has no mark</summary>
    public const double NoMarkClearance = 6;

    // in pixels of being pushed out: a square pixel of overlap, a pixel of a side through the
    // box, the middle of the box outside the angle, a turn all the way to a side
    const double OverlapCost = 10;
    const double SideCost = 1;
    const double OutsideCost = 25;
    const double TurnCost = 24;

    /// <summary>The farthest a number turns off the bisector, as a fraction of the half angle</summary>
    const double MaxTurn = 0.75;

    static readonly double[] Turns = [0, 0.25, -0.25, 0.5, -0.5, 0.75, -0.75];

    /// <summary>In pixels of being pushed out: what a place more than <see cref="Near"/> from where the number was must save to be taken</summary>
    public const double Hysteresis = 15;

    /// <summary>In pixels of being pushed out, per pixel: what the way from where the number was costs</summary>
    public const double Proximity = 1;

    /// <summary>In pixels: how far a number slides without the <see cref="Hysteresis"/></summary>
    public const double Near = 4;

    // the small steps a number slides by from where it was
    const double SlideTurn = 0.02;
    const double SlidePush = 1.5;

    /// <summary>A number's place relative to its angle: a turn off the bisector, a fraction of the half angle, and a push out in pixels</summary>
    public record struct Placement(double Turn, double Push);

    /// <summary>A number to place: its angle, its mark and the size of its box, in pixels but for the angles</summary>
    public class Wedge
    {
        public Point Vertex;

        /// <summary>The direction of the first side, in radians counterclockwise from the x axis, in the plane (y up)</summary>
        public double Start;

        /// <summary>The measure of the angle, 0 to 2π, which runs from the first side clockwise or counterclockwise</summary>
        public double Sweep;

        public bool IsClockwise;

        /// <summary>How far from the vertex the mark reaches: the box keeps that and <see cref="Gap"/> clear of it</summary>
        public double Clearance;

        public Size Size;

        /// <summary>Where the number was, if anywhere: it starts from there</summary>
        public Placement? Previous;
    }

    /// <summary>
    /// Puts the numbers of the angles at the vertex of this one that place themselves where
    /// they go now: their offsets are set, and those of the others moved on the canvas
    /// (the one asked about is placed by its own redraw, which asks this)
    /// </summary>
    public static void ArrangeAt(AngleMeasurement number)
    {
        var drawing = number.Drawing;
        var vertex = number.Dependencies[0];
        var numbers = vertex.Dependents
            .OfType<AngleMeasurement>()
            .Where(other => other == number || other.TakesPartInLayout(vertex))
            .Distinct()
            .OrderBy(other => drawing.Figures.IndexOf(other))
            .ToList();
        var placing = numbers.Where(other => other.PlacesItself).ToList();
        if (placing.IsEmpty())
        {
            return;
        }

        var corner = drawing.CoordinateSystem.ToPhysical(number.Point(0));
        var obstacles = numbers
            .Where(other => !other.PlacesItself)
            .Select(other => new Rect(corner + other.Offset, other.LayoutSize()))
            .ToList();
        var places = Arrange(placing.Select(other => other.LayoutWedge()).ToList(), obstacles, SidesOf(numbers));
        for (int i = 0; i < placing.Count; i++)
        {
            placing[i].LayoutPlacement = places[i].Placement;
            placing[i].Offset = places[i].TopLeft - corner;
            if (placing[i] != number)
            {
                placing[i].MoveToOffset();
            }
        }
    }

    /// <summary>The directions of the sides of the angles, out of their vertex, in radians in the plane</summary>
    public static List<double> SidesOf(IEnumerable<AngleMeasurement> numbers)
    {
        var sides = new List<double>();
        foreach (var number in numbers)
        {
            sides.Add(Math.GetAngle(number.Point(0), number.Point(1)));
            sides.Add(Math.GetAngle(number.Point(0), number.Point(2)));
        }

        return sides;
    }

    /// <summary>
    /// Where the boxes of the wedges go, in their order (top-left corners, and the places to
    /// remember), kept clear of the obstacles (boxes that stay where they are) and, where they
    /// can be, of the sides: the directions out of the vertex, in radians in the plane, that
    /// lines go along
    /// </summary>
    public static (Point TopLeft, Placement Placement)[] Arrange(IReadOnlyList<Wedge> wedges, IReadOnlyList<Rect> obstacles, IReadOnlyList<double> sides)
    {
        // every number where it was, for the first pass to keep clear of
        var boxes = new Rect?[wedges.Count];
        var placements = new Placement[wedges.Count];
        for (int i = 0; i < wedges.Count; i++)
        {
            if (wedges[i].Previous is Placement previous)
            {
                boxes[i] = BoxAt(wedges[i], previous);
            }
        }

        for (int pass = 0; pass < 2; pass++)
        {
            for (int i = 0; i < wedges.Count; i++)
            {
                var others = new List<Rect>(obstacles);
                for (int j = 0; j < boxes.Length; j++)
                {
                    if (j != i && boxes[j] != null)
                    {
                        others.Add(boxes[j].Value);
                    }
                }

                (boxes[i], placements[i]) = Place(wedges[i], others, sides);
            }
        }

        return boxes.Select((box, i) => (box.Value.TopLeft, placements[i])).ToArray();
    }

    static (Rect Box, Placement Placement) Place(Wedge wedge, List<Rect> others, IReadOnlyList<double> sides)
    {
        double preferred = Clearance(wedge, Direction(Bisector(wedge)));
        Point? previousMiddle = wedge.Previous is Placement previous ? BoxAt(wedge, previous).Center : null;
        var best = default(Rect);
        var bestPlacement = default(Placement);
        double bestCost = double.PositiveInfinity;

        // Whether the place is clean: nothing over it, inside the angle
        bool Consider(Point direction, double clearance, Placement placement)
        {
            double distance = clearance + placement.Push;
            var box = BoxAt(wedge, direction, distance);
            double overlap = Overlap(box, others);
            double across = sides.Sum(side => Chord(wedge.Vertex, Direction(side), box));
            bool inside = IsMiddleInside(wedge, box);
            double cost = overlap * OverlapCost
                + across * SideCost
                + (inside ? 0 : OutsideCost)
                + System.Math.Abs(placement.Turn) * TurnCost
                + System.Math.Max(0, distance - preferred);
            if (previousMiddle != null)
            {
                double way = box.Center.Distance(previousMiddle.Value);
                cost += way * Proximity + (way > Near ? Hysteresis : 0);
            }

            if (cost < bestCost)
            {
                bestCost = cost;
                best = box;
                bestPlacement = placement;
            }

            return overlap == 0 && across == 0 && inside;
        }

        foreach (var turn in Turns)
        {
            var direction = Direction(Bisector(wedge) + turn * wedge.Sweep / 2);
            double clearance = Clearance(wedge, direction);
            for (double push = 0; push <= MaxPush; push += PushStep)
            {
                // further out only costs more
                if (Consider(direction, clearance, new Placement(turn, push)))
                {
                    break;
                }
            }
        }

        // where it was, and a small step in every way from there
        if (wedge.Previous is Placement from)
        {
            foreach (var turnStep in new[] { 0, -SlideTurn, SlideTurn })
            {
                double turn = System.Math.Clamp(from.Turn + turnStep, -MaxTurn, MaxTurn);
                var direction = Direction(Bisector(wedge) + turn * wedge.Sweep / 2);
                double clearance = Clearance(wedge, direction);
                foreach (var pushStep in new[] { 0, -SlidePush, SlidePush })
                {
                    Consider(direction, clearance, new Placement(turn, System.Math.Clamp(from.Push + pushStep, 0, MaxPush)));
                }
            }
        }

        return (best, bestPlacement);
    }

    static double Bisector(Wedge wedge)
    {
        return wedge.Start + (wedge.IsClockwise ? -wedge.Sweep : wedge.Sweep) / 2;
    }

    static Rect BoxAt(Wedge wedge, Placement placement)
    {
        var direction = Direction(Bisector(wedge) + placement.Turn * wedge.Sweep / 2);
        return BoxAt(wedge, direction, Clearance(wedge, direction) + placement.Push);
    }

    /// <summary>How far out along the direction the middle of the box goes at the least: where the box clears the mark</summary>
    static double Clearance(Wedge wedge, Point direction)
    {
        // nearer the vertex than the mark at first, out of reach of it at the last
        double clearance = wedge.Clearance + Gap;
        double low = 0;
        double high = clearance + wedge.Size.Width + wedge.Size.Height;
        for (int i = 0; i < 30; i++)
        {
            double middle = (low + high) / 2;
            if (DistanceToBox(wedge.Vertex, BoxAt(wedge, direction, middle)) >= clearance)
            {
                high = middle;
            }
            else
            {
                low = middle;
            }
        }

        return high;
    }

    /// <summary>A unit vector in pixels (y down) along a direction in the plane (y up)</summary>
    static Point Direction(double angle)
    {
        return new Point(System.Math.Cos(angle), -System.Math.Sin(angle));
    }

    static Rect BoxAt(Wedge wedge, Point direction, double distance)
    {
        var middle = wedge.Vertex + direction * distance;
        return new Rect(
            middle.X - wedge.Size.Width / 2,
            middle.Y - wedge.Size.Height / 2,
            wedge.Size.Width,
            wedge.Size.Height);
    }

    /// <summary>Whether the middle of the box is inside the angle</summary>
    static bool IsMiddleInside(Wedge wedge, Rect box)
    {
        var offset = box.Center - wedge.Vertex;
        double angle = System.Math.Atan2(-offset.Y, offset.X);
        return Normalize(wedge.IsClockwise ? wedge.Start - angle : angle - wedge.Start) <= wedge.Sweep;
    }

    static double Normalize(double angle)
    {
        double turn = 2 * System.Math.PI;
        angle %= turn;
        return angle < 0 ? angle + turn : angle;
    }

    /// <summary>In pixels: how much of the ray from the origin along the direction (a unit vector) runs through the box</summary>
    static double Chord(Point origin, Point direction, Rect box)
    {
        double near = 0;
        double far = double.PositiveInfinity;
        if (!Slab(origin.X, direction.X, box.Left, box.Right, ref near, ref far)
            || !Slab(origin.Y, direction.Y, box.Top, box.Bottom, ref near, ref far))
        {
            return 0;
        }

        return far - near;
    }

    static bool Slab(double origin, double direction, double low, double high, ref double near, ref double far)
    {
        if (System.Math.Abs(direction) < 1e-12)
        {
            return origin >= low && origin <= high;
        }

        double first = (low - origin) / direction;
        double second = (high - origin) / direction;
        near = System.Math.Max(near, System.Math.Min(first, second));
        far = System.Math.Min(far, System.Math.Max(first, second));
        return near <= far;
    }

    static double DistanceToBox(Point point, Rect box)
    {
        double x = System.Math.Max(System.Math.Max(box.Left - point.X, point.X - box.Right), 0);
        double y = System.Math.Max(System.Math.Max(box.Top - point.Y, point.Y - box.Bottom), 0);
        return System.Math.Sqrt(x * x + y * y);
    }

    /// <summary>The area the box overlaps the others by, each with <see cref="Spacing"/> around it</summary>
    static double Overlap(Rect box, List<Rect> others)
    {
        double sum = 0;
        foreach (var other in others)
        {
            double width = System.Math.Min(box.Right, other.Right + Spacing) - System.Math.Max(box.Left, other.Left - Spacing);
            double height = System.Math.Min(box.Bottom, other.Bottom + Spacing) - System.Math.Max(box.Top, other.Top - Spacing);
            if (width > 0 && height > 0)
            {
                sum += width * height;
            }
        }

        return sum;
    }
}
