namespace DynamicGeometry
{
    public class PerpendicularLine : LineTwoPoints
    {
        public override PointPair Coordinates
        {
            get
            {
                var line = Dependencies.Line(0);
                if (Flipped) line = new PointPair(line.P2,line.P1);
                var point = Point(1);
                var coordinates = Math.GetPerpendicularLine(line, point);
                return coordinates;
            }
        }
    }
}