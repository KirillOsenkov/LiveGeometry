using Avalonia;

namespace DynamicGeometry
{
    public interface ILengthProvider : IFigure
    {
        double Length { get; }
    }

    public class DistanceMeasurement : Measurement, ILengthProvider
    {
        public override Point Anchor
        {
            get
            {
                return Midpoint();
            }
        }

        protected override string Kind
        {
            get
            {
                return "Distance";
            }
        }

        /// <summary>"AB": "Distance AB"</summary>
        public override string Construction
        {
            get
            {
                if (Dependencies.Count > 1 && Dependencies[0] is IPoint && Dependencies[1] is IPoint)
                {
                    return ConstructionText.Points(Dependencies[0], Dependencies[1]);
                }

                return Dependencies.Count > 0 ? ConstructionText.Length(Dependencies[0]) : null;
            }
        }

        /// <summary>
        /// A figure with a length, or two points. Anything else has nothing to measure (a
        /// segment that was replaced by a ray, a file that says so): the measurement then
        /// doesn't exist, where it used to throw on every redraw.
        /// </summary>
        bool HasSomethingToMeasure
        {
            get
            {
                return Dependencies.Count > 0
                    && (Dependencies[0] is ILengthProvider
                        || Dependencies.Count > 1 && Dependencies[0] is IPoint && Dependencies[1] is IPoint);
            }
        }

        public override void UpdateExistence()
        {
            base.UpdateExistence();
            if (Exists && !HasSomethingToMeasure)
            {
                Exists = false;
            }
        }

        public Point Midpoint()
        {
            if (!HasSomethingToMeasure)
            {
                return Math.InfinitePoint;
            }

            if (Dependencies[0] is ILengthProvider)
            {
                return Dependencies[0].Center;
            }
            return Math.Midpoint(Point(0), Point(1));
        }

        public double Distance
        {
            get
            {
                if (!HasSomethingToMeasure)
                {
                    return double.NaN;
                }

                if (Dependencies[0] is ILengthProvider)
                {
                    return (Dependencies[0] as ILengthProvider).Length;
                }
                return Point(0).Distance(Point(1));
            }
        }

        public double Length
        {
            get
            {
                return Distance;
            }
        }

        public override void UpdateVisual()
        {
            if (!HasSomethingToMeasure)
            {
                return;
            }

            base.UpdateVisual();
            Text = Math.Round(Distance, DecimalsToShow).ToString();
        }
    }
}
