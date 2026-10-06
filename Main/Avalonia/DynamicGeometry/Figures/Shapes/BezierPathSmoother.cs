using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace DynamicGeometry;

/// <summary>How a Bezier path works out the handles left to it (automatic ones) from where its anchors are</summary>
public enum BezierPathSmoothing
{
    /// <summary>On their anchors: the path goes from anchor to anchor in straight lines</summary>
    None,

    /// <summary>
    /// Hobby's (METAFONT's, MetaPost's): the angles at the anchors such that the curvature
    /// changes as little as it can across each one; four anchors on a circle give a circle
    /// </summary>
    Hobby,

    /// <summary>Centripetal Catmull-Rom: each anchor's handles along the line through its neighbors, from those alone</summary>
    [PropertyGridName("Catmull-Rom")]
    CatmullRom,

    /// <summary>The natural cubic spline (chord length): the curvature doesn't jump at all, which the whole path takes part in</summary>
    [PropertyGridName("Natural spline")]
    NaturalSpline
}

/// <summary>
/// Works out the automatic handles of a Bezier path from its anchors. A handle the user has
/// set (dragged, made a point of the drawing) stays as it is, and the automatic ones fit
/// themselves to it: an automatic handle across the anchor from it continues in its direction
/// (the anchor stays smooth), and one next to a handle on its anchor (a corner) or at the end
/// of an open path starts as the free end of a curve does. Every method gives the same
/// curve, only bigger, smaller, turned or mirrored, for the anchors so moved. The tension
/// makes the automatic handles shorter (above 1) or longer (below 1).
/// </summary>
public static class BezierPathSmoother
{
    /// <summary>
    /// The offsets of the handles from their anchors: those given (non-null) as they are,
    /// the automatic ones (null) worked out. A handle that bends no piece (the outer ones of
    /// an open path's ends), or a piece of no length, leaves an automatic handle on its anchor.
    /// </summary>
    public static (Point[] Ins, Point[] Outs) Smooth(
        IReadOnlyList<Point> anchors,
        IReadOnlyList<Point?> ins,
        IReadOnlyList<Point?> outs,
        bool closed,
        BezierPathSmoothing smoothing,
        double tension)
    {
        var path = new Anchors(anchors, ins, outs, closed);
        if (smoothing != BezierPathSmoothing.None && path.PieceCount > 0 && tension > 0 && tension.IsValidValue())
        {
            switch (smoothing)
            {
                case BezierPathSmoothing.Hobby:
                    Hobby(path, tension);
                    break;
                case BezierPathSmoothing.CatmullRom:
                    CatmullRom(path, tension);
                    break;
                case BezierPathSmoothing.NaturalSpline:
                    NaturalSpline(path, tension);
                    break;
            }
        }

        return (path.ResultIns, path.ResultOuts);
    }

    /// <summary>The anchors, the pieces between them and what is known at each anchor</summary>
    class Anchors
    {
        public Anchors(IReadOnlyList<Point> points, IReadOnlyList<Point?> ins, IReadOnlyList<Point?> outs, bool closed)
        {
            Points = points;
            Ins = ins;
            Outs = outs;
            Closed = closed;
            int count = points.Count;
            PieceCount = count < 2 ? 0 : closed ? count : count - 1;
            Chords = new Point[PieceCount];
            Lengths = new double[PieceCount];
            double size = points.Count == 0 ? 0 : points.Max(p => System.Math.Max(System.Math.Abs(p.X), System.Math.Abs(p.Y)));
            double tiny = 1e-12 * (1 + size);
            for (int i = 0; i < PieceCount; i++)
            {
                Chords[i] = points[(i + 1) % count].Minus(points[i]);
                Lengths[i] = Length(Chords[i]);
                if (!(Lengths[i] > tiny))
                {
                    Lengths[i] = 0;
                }
            }

            ResultIns = new Point[count];
            ResultOuts = new Point[count];
            for (int i = 0; i < count; i++)
            {
                ResultIns[i] = ins[i] ?? default;
                ResultOuts[i] = outs[i] ?? default;
            }
        }

        public IReadOnlyList<Point> Points { get; }

        public IReadOnlyList<Point?> Ins { get; }

        public IReadOnlyList<Point?> Outs { get; }

        public bool Closed { get; }

        public int Count => Points.Count;

        public int PieceCount { get; }

        /// <summary>From each anchor to the next</summary>
        public Point[] Chords { get; }

        /// <summary>The lengths of the chords; 0 for a piece too short to have a direction</summary>
        public double[] Lengths { get; }

        public Point[] ResultIns { get; }

        public Point[] ResultOuts { get; }

        /// <summary>The piece that comes to the anchor; -1 at the start of an open path</summary>
        public int PieceBefore(int anchor)
        {
            return Closed ? (anchor + Count - 1) % Count : anchor - 1;
        }

        /// <summary>The piece that leaves the anchor; -1 at the end of an open path</summary>
        public int PieceAfter(int anchor)
        {
            return anchor < PieceCount ? anchor : -1;
        }

        public int Next(int anchor)
        {
            return (anchor + 1) % Count;
        }

        public bool HasLength(int piece)
        {
            return piece >= 0 && Lengths[piece] > 0;
        }

        /// <summary>Whether both handles of the anchor are automatic and both pieces there have a length: the angle there is worked out</summary>
        public bool IsSmooth(int anchor)
        {
            return Ins[anchor] == null
                && Outs[anchor] == null
                && HasLength(PieceBefore(anchor))
                && HasLength(PieceAfter(anchor));
        }

        /// <summary>
        /// Where the path leaves the anchor, when that is given: by its out handle, or for an
        /// automatic one, straight on from the in handle. Null: free (a handle on its anchor,
        /// an end).
        /// </summary>
        public Point? OutDirection(int anchor)
        {
            if (Outs[anchor] is Point outOffset)
            {
                return Unit(outOffset);
            }

            if (PieceBefore(anchor) >= 0 && Ins[anchor] is Point inOffset)
            {
                return Unit(inOffset.Minus());
            }

            return null;
        }

        /// <summary>Where the path comes to the anchor going, when that is given</summary>
        public Point? InDirection(int anchor)
        {
            if (Ins[anchor] is Point inOffset)
            {
                return Unit(inOffset.Minus());
            }

            if (PieceAfter(anchor) >= 0 && Outs[anchor] is Point outOffset)
            {
                return Unit(outOffset);
            }

            return null;
        }

        /// <summary>
        /// The stretches of pieces between anchors where the angle is given or free (in the
        /// order of the path), each with the smooth anchors inside it; none when every anchor
        /// of a closed path is smooth (<see cref="AllSmooth"/>)
        /// </summary>
        public List<List<int>> Runs()
        {
            var runs = new List<List<int>>();
            for (int start = 0; start < Count; start++)
            {
                if (IsSmooth(start) || !HasLength(PieceAfter(start)))
                {
                    continue;
                }

                // the anchors of the run, the first and the last not smooth
                var run = new List<int>() { start };
                int anchor = start;
                do
                {
                    anchor = Next(anchor);
                    run.Add(anchor);
                }
                while (IsSmooth(anchor));
                runs.Add(run);
            }

            return runs;
        }

        public bool AllSmooth
        {
            get
            {
                return Closed && Enumerable.Range(0, Count).All(IsSmooth);
            }
        }

        /// <summary>How far the path turns at the anchor, from the chord before to the chord after</summary>
        public double Turn(int anchor)
        {
            return Angle(Chords[PieceBefore(anchor)], Chords[PieceAfter(anchor)]);
        }
    }

    #region Hobby

    // The angles a piece leaves its first anchor and comes to its second at, from its chord:
    // theta counterclockwise, phi clockwise (an arc bulging to the left has both positive).
    // At a smooth anchor the path doesn't turn: phi there + theta there = -(the turn of the
    // chords). Hobby's "mock curvature" (the curvature of a piece as far as it depends on the
    // angles in the first order: at its start 2 (theta + phi - 3 tension theta) tension / d,
    // at its end the same with phi for theta) is the same on both sides of each smooth
    // anchor, a linear equation per anchor. A free end has the curvature of the piece's other
    // end ("curl" 1).
    const double Curl = 1;

    // below this tension the equations lose their footing (MetaPost asks for 3/4 at least);
    // the handles are still made longer by the tension asked for
    const double LeastTension = 0.75;

    static void Hobby(Anchors path, double tension)
    {
        double c = 3 * System.Math.Max(tension, LeastTension) - 1;
        var theta = new double[path.PieceCount];
        var phi = new double[path.PieceCount];
        if (path.AllSmooth)
        {
            HobbyCycle(path, c, theta, phi);
        }
        else
        {
            foreach (var run in path.Runs())
            {
                HobbyRun(path, run, c, theta, phi);
            }
        }

        for (int piece = 0; piece < path.PieceCount; piece++)
        {
            if (!path.HasLength(piece))
            {
                continue;
            }

            int next = path.Next(piece);
            var chord = path.Chords[piece];
            if (path.Outs[piece] == null)
            {
                path.ResultOuts[piece] = Rotate(chord, theta[piece]).Scale(Velocity(theta[piece], phi[piece]) / (3 * tension));
            }

            if (path.Ins[next] == null)
            {
                path.ResultIns[next] = Rotate(chord, -phi[piece]).Scale(-Velocity(phi[piece], theta[piece]) / (3 * tension));
            }
        }
    }

    /// <summary>A closed path smooth at every anchor: an equation per anchor, round and round</summary>
    static void HobbyCycle(Anchors path, double c, double[] theta, double[] phi)
    {
        int count = path.Count;
        var sub = new double[count];
        var diagonal = new double[count];
        var super = new double[count];
        var right = new double[count];
        for (int k = 0; k < count; k++)
        {
            int before = path.PieceBefore(k);
            int next = path.Next(k);
            double dBefore = path.Lengths[before];
            double dAfter = path.Lengths[k];
            sub[k] = 1 / dBefore;
            diagonal[k] = c * (1 / dBefore + 1 / dAfter);
            super[k] = 1 / dAfter;
            right[k] = -c * path.Turn(k) / dBefore - path.Turn(next) / dAfter;
        }

        var x = SolveCyclic(sub, diagonal, super, right);
        if (x == null)
        {
            return;
        }

        for (int k = 0; k < count; k++)
        {
            int next = path.Next(k);
            theta[k] = x[k];
            phi[k] = -path.Turn(next) - x[next];
        }
    }

    /// <summary>
    /// The pieces from one anchor that isn't smooth to the next: the angles at the smooth
    /// anchors between them, and at the two ends, given or free
    /// </summary>
    static void HobbyRun(Anchors path, List<int> run, double c, double[] theta, double[] phi)
    {
        // the unknowns: theta at the start of each piece, then phi at the end of the last
        int m = run.Count - 1;
        var pieces = run.Take(m).ToList();
        double D(int j) => path.Lengths[pieces[j]];
        double Psi(int j) => path.Turn(run[j]);

        var start = path.OutDirection(run[0]);
        var end = path.InDirection(run[m]);
        if (m == 1 && start == null && end == null)
        {
            // two free ends: a straight piece
            theta[pieces[0]] = 0;
            phi[pieces[0]] = 0;
            return;
        }

        var sub = new double[m + 1];
        var diagonal = new double[m + 1];
        var super = new double[m + 1];
        var right = new double[m + 1];
        if (start is Point startDirection)
        {
            diagonal[0] = 1;
            right[0] = Angle(path.Chords[pieces[0]], startDirection);
        }
        else
        {
            // curl: theta0 (c + curl) = phi1 (1 + curl c), phi1 = -psi1 - theta1 unless it is the end
            diagonal[0] = c + Curl;
            if (m == 1)
            {
                super[0] = -(1 + Curl * c);
            }
            else
            {
                super[0] = 1 + Curl * c;
                right[0] = -(1 + Curl * c) * Psi(1);
            }
        }

        for (int j = 1; j < m; j++)
        {
            sub[j] = 1 / D(j - 1);
            diagonal[j] = c * (1 / D(j - 1) + 1 / D(j));
            right[j] = -c * Psi(j) / D(j - 1);
            if (j + 1 < m)
            {
                super[j] = 1 / D(j);
                right[j] -= Psi(j + 1) / D(j);
            }
            else
            {
                // the next unknown is phi at the end itself
                super[j] = -1 / D(j);
            }
        }

        if (end is Point endDirection)
        {
            diagonal[m] = 1;
            right[m] = Angle(endDirection, path.Chords[pieces[m - 1]]);
        }
        else
        {
            // curl: phi (c + curl) = theta of the last piece (1 + curl c)
            sub[m] = -(1 + Curl * c);
            diagonal[m] = c + Curl;
        }

        var x = SolveTridiagonal(sub, diagonal, super, right);
        if (x == null)
        {
            return;
        }

        for (int j = 0; j < m; j++)
        {
            theta[pieces[j]] = x[j];
            phi[pieces[j]] = j + 1 < m ? -Psi(j + 1) - x[j + 1] : x[m];
        }
    }

    /// <summary>
    /// How long Hobby makes a handle, in thirds of the chord, for a piece that leaves at
    /// theta and comes in at phi (and the other handle with the two swapped): 1 for a
    /// straight piece, 1.1716 for a quarter of a circle (the circle's 0.5523 of the radius)
    /// </summary>
    static double Velocity(double theta, double phi)
    {
        double sinTheta = System.Math.Sin(theta);
        double cosTheta = System.Math.Cos(theta);
        double sinPhi = System.Math.Sin(phi);
        double cosPhi = System.Math.Cos(phi);
        double sqrt5 = System.Math.Sqrt(5);
        double numerator = 2 + System.Math.Sqrt(2) * (sinTheta - sinPhi / 16) * (sinPhi - sinTheta / 16) * (cosTheta - cosPhi);
        double denominator = 1 + (sqrt5 - 1) / 2 * cosTheta + (3 - sqrt5) / 2 * cosPhi;

        // (METAFONT's limit: a piece that turns almost all the way round)
        const double longest = 4;
        return denominator > numerator / longest ? numerator / denominator : longest;
    }

    #endregion

    #region Catmull-Rom

    // The tangent at an anchor from its two neighbors, the knots spaced by the square roots
    // of the chords (centripetal: no loops or cusps inside a piece), the handles a third of
    // the tangent times the piece's spacing. Where there is only one neighbor to go by (an
    // end, a handle on its anchor next to it) the handle points halfway to the handle at the
    // other end of its piece: the curve doesn't bend at the free end.
    static void CatmullRom(Anchors path, double tension)
    {
        int count = path.Count;
        var inKnown = new bool[count];
        var outKnown = new bool[count];
        for (int k = 0; k < count; k++)
        {
            inKnown[k] = path.Ins[k] != null;
            outKnown[k] = path.Outs[k] != null;
            int before = path.PieceBefore(k);
            int after = path.PieceAfter(k);
            if (!path.HasLength(before) || !path.HasLength(after))
            {
                continue;
            }

            var point = path.Points[k];
            var previous = path.Points[(k + count - 1) % count];
            var next = path.Points[path.Next(k)];
            double spacingBefore = System.Math.Sqrt(path.Lengths[before]);
            double spacingAfter = System.Math.Sqrt(path.Lengths[after]);
            var tangent = point.Minus(previous).Scale(1 / spacingBefore)
                .Minus(next.Minus(previous).Scale(1 / (spacingBefore + spacingAfter)))
                .Plus(next.Minus(point).Scale(1 / spacingAfter));
            var outOffset = tangent.Scale(spacingAfter / 3);
            var inOffset = tangent.Scale(-spacingBefore / 3);
            if (path.Outs[k] == null && path.Ins[k] == null)
            {
                path.ResultOuts[k] = outOffset;
                path.ResultIns[k] = inOffset;
                outKnown[k] = inKnown[k] = true;
            }
            else if (path.Outs[k] == null && path.OutDirection(k) is Point outDirection)
            {
                path.ResultOuts[k] = outDirection.Scale(Length(outOffset));
                outKnown[k] = true;
            }
            else if (path.Ins[k] == null && path.InDirection(k) is Point inDirection)
            {
                path.ResultIns[k] = inDirection.Scale(-Length(inOffset));
                inKnown[k] = true;
            }
        }

        // the free ends, halfway to the handle at the other end of the piece (both free: a
        // straight piece, the handles at its thirds)
        for (int piece = 0; piece < path.PieceCount; piece++)
        {
            if (!path.HasLength(piece))
            {
                continue;
            }

            int first = piece;
            int second = path.Next(piece);
            var from = path.Points[first];
            var to = path.Points[second];
            if (!outKnown[first] && !inKnown[second])
            {
                path.ResultOuts[first] = path.Chords[piece].Scale(1.0 / 3);
                path.ResultIns[second] = path.Chords[piece].Scale(-1.0 / 3);
            }
            else if (!outKnown[first])
            {
                path.ResultOuts[first] = to.Plus(path.ResultIns[second]).Minus(from).Scale(0.5);
            }
            else if (!inKnown[second])
            {
                path.ResultIns[second] = from.Plus(path.ResultOuts[first]).Minus(to).Scale(0.5);
            }
        }

        ScaleAutomatic(path, 1 / tension);
    }

    #endregion

    #region Natural spline

    // The pieces as one curve with continuous first and second derivatives, the parameter
    // running along each piece as far as its chord is long (the derivative is about a unit
    // vector then): D at each anchor, the out handle D d / 3, the in handle -D d / 3 with the
    // chord before. A free end has no curvature; a given direction is D itself.
    static void NaturalSpline(Anchors path, double tension)
    {
        if (path.AllSmooth)
        {
            NaturalCycle(path);
        }
        else
        {
            foreach (var run in path.Runs())
            {
                NaturalRun(path, run);
            }
        }

        ScaleAutomatic(path, 1 / tension);
    }

    static void NaturalCycle(Anchors path)
    {
        int count = path.Count;
        var sub = new double[count];
        var diagonal = new double[count];
        var super = new double[count];
        var rightX = new double[count];
        var rightY = new double[count];
        for (int k = 0; k < count; k++)
        {
            int before = path.PieceBefore(k);
            double hBefore = path.Lengths[before];
            double hAfter = path.Lengths[k];
            sub[k] = 1 / hBefore;
            diagonal[k] = 2 * (1 / hBefore + 1 / hAfter);
            super[k] = 1 / hAfter;
            var right = path.Chords[before].Scale(3 / (hBefore * hBefore)).Plus(path.Chords[k].Scale(3 / (hAfter * hAfter)));
            rightX[k] = right.X;
            rightY[k] = right.Y;
        }

        var x = SolveCyclic(sub, diagonal, super, rightX);
        var y = SolveCyclic(sub, diagonal, super, rightY);
        if (x == null || y == null)
        {
            return;
        }

        for (int k = 0; k < count; k++)
        {
            var derivative = new Point(x[k], y[k]);
            path.ResultOuts[k] = derivative.Scale(path.Lengths[k] / 3);
            path.ResultIns[k] = derivative.Scale(-path.Lengths[path.PieceBefore(k)] / 3);
        }
    }

    static void NaturalRun(Anchors path, List<int> run)
    {
        int m = run.Count - 1;
        var pieces = run.Take(m).ToList();
        double H(int j) => path.Lengths[pieces[j]];
        Point Chord(int j) => path.Chords[pieces[j]];

        var sub = new double[m + 1];
        var diagonal = new double[m + 1];
        var super = new double[m + 1];
        var right = new Point[m + 1];
        if (path.OutDirection(run[0]) is Point start)
        {
            diagonal[0] = 1;
            right[0] = start;
        }
        else
        {
            diagonal[0] = 2;
            super[0] = 1;
            right[0] = Chord(0).Scale(3 / H(0));
        }

        for (int j = 1; j < m; j++)
        {
            sub[j] = 1 / H(j - 1);
            diagonal[j] = 2 * (1 / H(j - 1) + 1 / H(j));
            super[j] = 1 / H(j);
            right[j] = Chord(j - 1).Scale(3 / (H(j - 1) * H(j - 1))).Plus(Chord(j).Scale(3 / (H(j) * H(j))));
        }

        if (path.InDirection(run[m]) is Point end)
        {
            diagonal[m] = 1;
            right[m] = end;
        }
        else
        {
            sub[m] = 1;
            diagonal[m] = 2;
            right[m] = Chord(m - 1).Scale(3 / H(m - 1));
        }

        var x = SolveTridiagonal(sub, diagonal, super, right.Select(p => p.X).ToArray());
        var y = SolveTridiagonal(sub, diagonal, super, right.Select(p => p.Y).ToArray());
        if (x == null || y == null)
        {
            return;
        }

        for (int j = 0; j < m; j++)
        {
            int first = run[j];
            int second = run[j + 1];
            if (path.Outs[first] == null)
            {
                path.ResultOuts[first] = new Point(x[j], y[j]).Scale(H(j) / 3);
            }

            if (path.Ins[second] == null)
            {
                path.ResultIns[second] = new Point(x[j + 1], y[j + 1]).Scale(-H(j) / 3);
            }
        }
    }

    #endregion

    #region Helpers

    /// <summary>The automatic handles longer or shorter by the factor (the tension of the methods that have none of their own)</summary>
    static void ScaleAutomatic(Anchors path, double factor)
    {
        for (int k = 0; k < path.Count; k++)
        {
            if (path.Ins[k] == null)
            {
                path.ResultIns[k] = path.ResultIns[k].Scale(factor);
            }

            if (path.Outs[k] == null)
            {
                path.ResultOuts[k] = path.ResultOuts[k].Scale(factor);
            }
        }
    }

    static double Length(Point vector)
    {
        return System.Math.Sqrt(vector.X * vector.X + vector.Y * vector.Y);
    }

    static Point? Unit(Point vector)
    {
        double length = Length(vector);
        return length > 0 && length.IsValidValue() ? vector.Scale(1 / length) : null;
    }

    static Point Rotate(Point vector, double angle)
    {
        double cos = System.Math.Cos(angle);
        double sin = System.Math.Sin(angle);
        return new Point(vector.X * cos - vector.Y * sin, vector.X * sin + vector.Y * cos);
    }

    /// <summary>The angle counterclockwise from one direction to the other, in (-pi, pi]</summary>
    static double Angle(Point from, Point to)
    {
        double angle = System.Math.Atan2(from.X * to.Y - from.Y * to.X, from.X * to.X + from.Y * to.Y);
        return angle <= -System.Math.PI ? System.Math.PI : angle;
    }

    /// <summary>
    /// Thomas's algorithm: row i says sub[i] x[i-1] + diagonal[i] x[i] + super[i] x[i+1] =
    /// right[i]; null when it breaks down
    /// </summary>
    static double[] SolveTridiagonal(double[] sub, double[] diagonal, double[] super, double[] right)
    {
        int count = diagonal.Length;
        var superPrime = new double[count];
        var rightPrime = new double[count];
        for (int i = 0; i < count; i++)
        {
            double pivot = diagonal[i] - (i > 0 ? sub[i] * superPrime[i - 1] : 0);
            if (System.Math.Abs(pivot) < 1e-300)
            {
                return null;
            }

            superPrime[i] = super[i] / pivot;
            rightPrime[i] = (right[i] - (i > 0 ? sub[i] * rightPrime[i - 1] : 0)) / pivot;
        }

        var x = new double[count];
        for (int i = count - 1; i >= 0; i--)
        {
            x[i] = rightPrime[i] - (i + 1 < count ? superPrime[i] * x[i + 1] : 0);
        }

        return x.All(value => value.IsValidValue()) ? x : null;
    }

    /// <summary>
    /// The same round a cycle: sub[0] goes with the last unknown and super[last] with the
    /// first (Sherman and Morrison's correction of the tridiagonal solution)
    /// </summary>
    static double[] SolveCyclic(double[] sub, double[] diagonal, double[] super, double[] right)
    {
        int count = diagonal.Length;
        if (count == 2)
        {
            double a = diagonal[0];
            double b = sub[0] + super[0];
            double c = sub[1] + super[1];
            double d = diagonal[1];
            double determinant = a * d - b * c;
            if (System.Math.Abs(determinant) < 1e-300)
            {
                return null;
            }

            return new[] { (right[0] * d - b * right[1]) / determinant, (a * right[1] - c * right[0]) / determinant };
        }

        double corner = super[count - 1];
        double otherCorner = sub[0];
        double gamma = -diagonal[0];
        var changed = (double[])diagonal.Clone();
        changed[0] -= gamma;
        changed[count - 1] -= corner * otherCorner / gamma;
        var plainSub = (double[])sub.Clone();
        var plainSuper = (double[])super.Clone();
        plainSub[0] = 0;
        plainSuper[count - 1] = 0;
        var x = SolveTridiagonal(plainSub, changed, plainSuper, right);
        var u = new double[count];
        u[0] = gamma;
        u[count - 1] = corner;
        var z = SolveTridiagonal(plainSub, changed, plainSuper, u);
        if (x == null || z == null)
        {
            return null;
        }

        double denominator = 1 + z[0] + otherCorner * z[count - 1] / gamma;
        if (System.Math.Abs(denominator) < 1e-300)
        {
            return null;
        }

        double factor = (x[0] + otherCorner * x[count - 1] / gamma) / denominator;
        for (int i = 0; i < count; i++)
        {
            x[i] -= factor * z[i];
        }

        return x;
    }

    #endregion
}
