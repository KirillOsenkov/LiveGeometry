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