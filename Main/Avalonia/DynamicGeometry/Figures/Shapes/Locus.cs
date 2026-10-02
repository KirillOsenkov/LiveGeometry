using System.Collections.Generic;
using Avalonia;

namespace DynamicGeometry
{
    public class Locus : Curve, ILinearFigure
    {
        private List<IFigure> figuresToRecalculate;

        /// <summary>
        /// What lies between the point that slides and the point traced, worked out anew
        /// each time: kept from the first time, it went on naming a figure that had since
        /// been replaced by another. No chain - the second point is not on a figure, or the
        /// first is not built on it - is a locus of nothing: no curve, rather than an
        /// exception from every move (PointSnapping keeps both points what they are; a
        /// file may still say anything).
        /// </summary>
        public override void Recalculate()
        {
            var point = mDependencies.Count > 0 ? mDependencies[0] as IPoint : null;
            var pointOnFigure = mDependencies.Count > 1 ? mDependencies[1] as PointOnFigure : null;
            figuresToRecalculate = point != null && pointOnFigure != null && point != pointOnFigure
                ? GetFiguresToRecalculate(pointOnFigure, point)
                : null;
        }

        protected override void OnDependenciesChanged()
        {
            figuresToRecalculate = null;
        }

        // the samples of the curve as GetPoints took them last: where the sliding point
        // was (its parameter on its figure) and where that put the traced point
        readonly List<double> sampleParameters = new List<double>();
        readonly List<Point> samplePoints = new List<Point>();

        bool HasChain
        {
            get { return figuresToRecalculate != null && figuresToRecalculate.Count > 0; }
        }

        /// <summary>Where the traced point is when the sliding point is at the parameter; the caller puts the sliding point back</summary>
        /// <returns>
        /// <see cref="Curve.Gap"/> where the traced point is not there (an intersection
        /// that the sliding point has taken apart): the last place it was at would be
        /// joined to the next one by a chord across the gap
        /// </returns>
        Point Trace(PointOnFigure sliding, IPoint traced, double parameter)
        {
            sliding.Parameter = parameter;
            for (int i = 0; i < figuresToRecalculate.Count; i++)
            {
                // (whether each is there, as a move of the sliding point by hand would ask)
                figuresToRecalculate[i].UpdateExistence();
                figuresToRecalculate[i].Recalculate();
            }

            return traced.Exists && traced.Coordinates.Exists() ? traced.Coordinates : Gap;
        }

        public override void GetPoints(List<Point> result)
        {
            sampleParameters.Clear();
            samplePoints.Clear();
            if (!HasChain)
            {
                return;
            }

            var point = mDependencies[0] as IPoint;
            var pointOnFigure = mDependencies[1] as PointOnFigure;
            var domain = pointOnFigure.LinearFigure.GetParameterDomain();
            var oldParameter = pointOnFigure.Parameter;

            if (Samples > 0)
            {
                // by the count, not by adding the step up: rounding decided whether the last
                // sample but one was taken, and the curve had 60 or 61 points from one move to
                // the next
                var steps = Samples;
                for (int i = 0; i <= steps; i++)
                {
                    double lambda = i == steps ? domain.Item2 : domain.Item1 + (domain.Item2 - domain.Item1) * i / steps;
                    sampleParameters.Add(lambda);
                    samplePoints.Add(Trace(pointOnFigure, point, lambda));
                }
            }
            else
            {
                SampleAdaptively(pointOnFigure, point, domain);
            }

            result.AddRange(samplePoints);
            Trace(pointOnFigure, point, oldParameter);
        }

        #region Samples

        /// <summary>
        /// A fixed number of even steps of the sliding point, or 0 for as many as it takes to
        /// look like a curve (<see cref="SampleAdaptively"/>). Only the Spiral says a number
        /// (60): its point is that its curve is that many straight pieces.
        /// </summary>
        public int Samples { get; set; }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            Samples = (int)System.Math.Max(0, System.Math.Min(MaxSamples, element.ReadDouble("Samples")));
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            if (Samples > 0)
            {
                writer.WriteAttributeString("Samples", Samples.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        /// <summary>Even steps of the sliding point that adaptive sampling starts from</summary>
        const int InitialSteps = 60;

        /// <summary>In pixels: how far the curve may stray from the straight piece drawn between two samples</summary>
        const double Tolerance = 0.5;

        /// <summary>In pixels: two samples still this far apart when halving the step can't bring them closer are a jump, not a piece of the curve</summary>
        const double JumpPixels = 20;

        /// <summary>How many times a step of the samples may be halved</summary>
        const int MaxDepth = 12;

        /// <summary>Samples a curve may take in all, counting those found not to be needed</summary>
        const int MaxSamples = 1000;

        /// <summary>
        /// Steps outward past an open end of the sliding point's figure, each twice as long
        /// as the last: from twice the window's stretch of a line to a thousand times that
        /// </summary>
        const int OpenEndSteps = 10;

        /// <summary>In sizes of the window: a traced point further than this from it is taken for a gap, not drawn</summary>
        const double FarReach = 20;

        /// <summary>A sample, and whether the piece from it to the next one is done (needs no more samples between)</summary>
        class Sample
        {
            public double Parameter;
            public Point Point;
            public bool Done;
        }

        /// <summary>
        /// <see cref="InitialSteps"/> even steps, then, round after round, every step halved
        /// where the curve strays from the straight piece between two samples by more than
        /// <see cref="Tolerance"/> pixels, or a gap begins or ends there - at most
        /// <see cref="MaxDepth"/> rounds and <see cref="MaxSamples"/> samples. Round after
        /// round rather than one stretch at a time, so that a curve that would take more
        /// (sin(1/x), wiggling without end) is drawn evenly coarse when the samples run out,
        /// not fine at its start and in straight pieces after. The image of a line in a
        /// circle, sampled evenly along the line, was a hexagon on its far side and missed
        /// the stretch near the center, which is where the line's far parts go: past the
        /// open ends of a line, a ray or a graph (<see cref="GetOpenEnds"/>) the steps go
        /// outward, each twice as long, and a traced point that runs away beyond
        /// <see cref="FarReach"/> sizes of the window is a gap rather than a coordinate too
        /// large to draw.
        /// </summary>
        void SampleAdaptively(PointOnFigure sliding, IPoint traced, Tuple<double, double> domain)
        {
            var coordinateSystem = Drawing.CoordinateSystem;
            double minX = coordinateSystem.MinimalVisibleX;
            double maxX = coordinateSystem.MaximalVisibleX;
            double minY = coordinateSystem.MinimalVisibleY;
            double maxY = coordinateSystem.MaximalVisibleY;
            var windowCenter = new Point((minX + maxX) / 2, (minY + maxY) / 2);
            double reach = FarReach * System.Math.Max(maxX - minX, maxY - minY);
            double unitLength = coordinateSystem.UnitLength;

            int samplesLeft = MaxSamples;
            Sample Take(double parameter)
            {
                samplesLeft--;
                var place = Trace(sliding, traced, parameter);
                return new Sample()
                {
                    Parameter = parameter,
                    Point = place.Exists() && place.Distance(windowCenter) <= reach ? place : Gap
                };
            }

            var parameters = new List<double>();
            GetOpenEnds(sliding.LinearFigure, out bool openLow, out bool openHigh);
            double span = domain.Item2 - domain.Item1;
            if (openLow)
            {
                for (int i = OpenEndSteps - 1; i >= 0; i--)
                {
                    parameters.Add(domain.Item1 - span * System.Math.Pow(2, i));
                }
            }

            for (int i = 0; i <= InitialSteps; i++)
            {
                parameters.Add(i == InitialSteps ? domain.Item2 : domain.Item1 + span * i / InitialSteps);
            }

            if (openHigh)
            {
                for (int i = 0; i < OpenEndSteps; i++)
                {
                    parameters.Add(domain.Item2 + span * System.Math.Pow(2, i));
                }
            }

            var samples = new List<Sample>(parameters.Count);
            foreach (var parameter in parameters)
            {
                samples.Add(Take(parameter));
            }

            // a round halves every piece that isn't done
            int round = 0;
            bool halved = true;
            while (halved && round < MaxDepth && samplesLeft > 0)
            {
                round++;
                halved = false;
                var next = new List<Sample>(samples.Count * 2);
                for (int i = 0; i < samples.Count; i++)
                {
                    var from = samples[i];
                    next.Add(from);
                    if (from.Done || i == samples.Count - 1 || samplesLeft <= 0)
                    {
                        continue;
                    }

                    var to = samples[i + 1];
                    bool fromExists = from.Point.Exists();
                    bool toExists = to.Point.Exists();
                    if (!fromExists && !toExists)
                    {
                        // nothing between two gaps, as far as the samples tell
                        from.Done = true;
                        continue;
                    }

                    var middle = Take((from.Parameter + to.Parameter) / 2);
                    if (fromExists && toExists && middle.Point.Exists()
                        && DistanceToSegment(middle.Point, from.Point, to.Point) * unitLength <= Tolerance)
                    {
                        from.Done = true;
                        continue;
                    }

                    next.Add(middle);
                    halved = true;
                }

                samples = next;
            }

            // Pieces still far apart after every round of halving are jumps (through
            // infinity, across a pole): not joined. Unless the samples ran out first - a
            // piece is long then because nobody looked into it.
            bool ranOut = samplesLeft <= 0 && halved;            for (int i = 0; i < samples.Count; i++)
            {
                var sample = samples[i];
                sampleParameters.Add(sample.Parameter);
                samplePoints.Add(sample.Point);
                if (!ranOut && !sample.Done && i < samples.Count - 1)
                {
                    var to = samples[i + 1];
                    if (sample.Point.Exists() && to.Point.Exists()
                        && sample.Point.Distance(to.Point) * unitLength > JumpPixels)
                    {
                        sampleParameters.Add((sample.Parameter + to.Parameter) / 2);
                        samplePoints.Add(Gap);
                    }
                }
            }
        }

        static double DistanceToSegment(Point point, Point from, Point to)
        {
            double dx = to.X - from.X;
            double dy = to.Y - from.Y;
            double lengthSquared = dx * dx + dy * dy;
            double ratio = lengthSquared > 0
                ? ((point.X - from.X) * dx + (point.Y - from.Y) * dy) / lengthSquared
                : 0;
            ratio = System.Math.Max(0, System.Math.Min(1, ratio));
            return point.Distance(new Point(from.X + dx * ratio, from.Y + dy * ratio));
        }

        /// <summary>
        /// Which ends of the parameters of the figure are only where the window cuts it, so
        /// that the figure goes on past them: both of a line's and a graph's, the far end of
        /// a ray's, a locus's as of its sliding point's figure. A segment, a circle or
        /// another curve ends where its parameters do.
        /// </summary>
        static void GetOpenEnds(ILinearFigure figure, out bool low, out bool high)
        {
            if (figure is FunctionGraph)
            {
                low = true;
                high = true;
            }
            else if (figure is LineBase && !(figure is Segment))
            {
                high = true;
                low = !(figure is Ray) || (figure is AngleBisector bisector && bisector.IsLine);
            }
            else if (figure is Locus locus && locus.mDependencies.Count > 1 && locus.mDependencies[1] is PointOnFigure sliding)
            {
                // a locus goes on as far as what its point slides along (the image of a
                // locus that is the image of a line missed the stretch near the center)
                GetOpenEnds(sliding.LinearFigure, out low, out high);
            }
            else
            {
                low = false;
                high = false;
            }
        }

        #endregion

        #region A point on the locus

        // A point on the locus is where the traced point is for some place of the sliding
        // point, and that place - the sliding point's parameter on its own figure - is
        // what the point keeps. (It kept a fraction of the length of the drawn curve. But
        // how much of the locus is drawn depends on the view when the sliding point is on
        // a line: the visible part of the line, twice over. The point then slid along the
        // curve with every pan and zoom, and a saved drawing opened with it elsewhere.)

        public override double GetNearestParameterFromPoint(Point point)
        {
            if (samplePoints.Count == 0)
            {
                return 0;
            }

            // the nearest place on the curve as drawn, and the parameter in proportion
            // between the two samples it lies between
            double nearest = double.MaxValue;
            double parameter = sampleParameters[0];
            for (int i = 0; i < samplePoints.Count - 1; i++)
            {
                var from = samplePoints[i];
                var to = samplePoints[i + 1];
                double dx = to.X - from.X;
                double dy = to.Y - from.Y;
                double lengthSquared = dx * dx + dy * dy;
                double ratio = lengthSquared > 0
                    ? ((point.X - from.X) * dx + (point.Y - from.Y) * dy) / lengthSquared
                    : 0;
                ratio = System.Math.Max(0, System.Math.Min(1, ratio));
                double x = from.X + dx * ratio - point.X;
                double y = from.Y + dy * ratio - point.Y;
                double distance = x * x + y * y;

                // (a sample where the traced point was nowhere compares as false)
                if (distance < nearest)
                {
                    nearest = distance;
                    parameter = sampleParameters[i] + (sampleParameters[i + 1] - sampleParameters[i]) * ratio;
                }
            }

            return parameter;
        }

        /// <summary>
        /// Worked out, not read off the drawn curve: the sliding point is put at the
        /// parameter, what lies between it and the traced point is recalculated, and
        /// everything is put back - as for every sample of the curve.
        /// </summary>
        public override Point GetPointFromParameter(double parameter)
        {
            var point = mDependencies.Count > 0 ? mDependencies[0] as IPoint : null;
            var pointOnFigure = mDependencies.Count > 1 ? mDependencies[1] as PointOnFigure : null;
            if (point == null || pointOnFigure == null || !HasChain || !parameter.IsValidValue())
            {
                return Math.InfinitePoint;
            }

            var oldParameter = pointOnFigure.Parameter;
            var result = Trace(pointOnFigure, point, parameter);
            Trace(pointOnFigure, point, oldParameter);
            return result.Exists() ? result : Math.InfinitePoint;
        }

        /// <summary>The parameters of the sliding point: those of the figure it slides on</summary>
        public override Tuple<double, double> GetParameterDomain()
        {
            var pointOnFigure = mDependencies.Count > 1 ? mDependencies[1] as PointOnFigure : null;
            return pointOnFigure != null ? pointOnFigure.LinearFigure.GetParameterDomain() : base.GetParameterDomain();
        }

        #endregion

        List<IFigure> GetFiguresToRecalculate(PointOnFigure pointOnFigure, IPoint dependentPoint)
        {
            return DependencyAlgorithms.FindImpactedDependencyChain(pointOnFigure, dependentPoint);
        }
    }
}
