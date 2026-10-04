namespace DynamicGeometry
{
    public class ParallelLine : LineTwoPoints
    {
        protected override string Kind
        {
            get
            {
                return "Parallel line";
            }
        }

        /// <summary>"to line AB through E"</summary>
        public override string Construction
        {
            get
            {
                if (Dependencies.Count < 2)
                {
                    return null;
                }

                return "to " + ConstructionText.Of(Dependencies[0]) + " through " + ConstructionText.Of(Dependencies[1]);
            }
        }

        protected override bool IsThroughTwoPoints
        {
            get
            {
                return false;
            }
        }

        public override PointPair Coordinates
        {
            get
            {
                PointPair coordinates;
                PointPair parentLine = Dependencies.Line(0);
                Avalonia.Point point = Point(1);

                coordinates = new PointPair()
                {
                    P1 = point,
                    P2 = point.Plus(parentLine.P2.Minus(parentLine.P1))
                };
                return coordinates;
            }
        }
    }
}