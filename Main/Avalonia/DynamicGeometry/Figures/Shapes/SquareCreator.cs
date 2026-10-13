using System.Collections.Generic;
using System.ComponentModel;
using Avalonia;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Shapes)]
    [Order(2)]
    public class SquareCreator : ShapeCreator
    {
        protected override IEnumerable<IFigure> CreateFigures()
        {
            var p1 = FoundDependencies[0] as IPoint;
            var p2 = FoundDependencies[1] as IPoint;
            if (p1.Coordinates.X == p2.Coordinates.X && p1.Coordinates.Y == p2.Coordinates.Y)
            {
                var p2AsMovable = p2 as IMovable;
                if (p2AsMovable != null)
                {
                    p2AsMovable.MoveTo(p2.Coordinates.Plus(new Point(.01, 0)));
                }
            }

            // a side drawn between the two points already (segment AB, then a square on A and
            // B) is the square's first side, as with the Triangle and Polygon tools: a second
            // segment lay on top of it as AB2
            var existingSide = Drawing.Figures.FindLine(p1, p2);
            var side0 = existingSide == null ? Factory.CreateSegment(Drawing, p1, p2) : null;
            var circle = Factory.CreateCircle(Drawing, new[] { p2, p1 });
            var perpendicular = Factory.CreatePerpendicularLine(Drawing, new IFigure[] { (IFigure)side0 ?? existingSide, p2 });
            // aim a counterclockwise turn from p1 -> p2, so that a square drawn left to right
            // stands on its first side; by the points as clicked, not by the perpendicular, whose
            // direction follows the side's: an existing side may run from p2 to p1 (the side of
            // a square drawn before), and the square came out on the other side of it
            var p1Coordinates = p1.Coordinates;
            var p2Coordinates = p2.Coordinates;
            var counterclockwise = new Point(
                p2Coordinates.X - (p2Coordinates.Y - p1Coordinates.Y),
                p2Coordinates.Y + (p2Coordinates.X - p1Coordinates.X));
            var intersection = Factory.CreateIntersectionPoint(Drawing, circle, perpendicular, counterclockwise);
            var midpoint = Factory.CreateMidPoint(Drawing, new IFigure[] { intersection, p1 });
            var reflectedPoint = Factory.CreateReflectedPoint(Drawing, new IFigure[] { p2, midpoint });
            var side1 = Factory.CreateSegment(Drawing, p2, intersection);
            var side2 = Factory.CreateSegment(Drawing, intersection, reflectedPoint);
            var side3 = Factory.CreateSegment(Drawing, reflectedPoint, p1);
            var polygon = Factory.CreatePolygon(Drawing, new IFigure[] { p1, p2, intersection, reflectedPoint });
            var added = new IFigure[]
            {
                side0,
                circle,
                perpendicular, 
                intersection, 
                midpoint,
                reflectedPoint,
                side1,
                side2,
                side3,
                polygon
            };

            circle.Visible = false;
            perpendicular.Visible = false;
            midpoint.Visible = false;

            foreach (var item in added)
            {
                if (item != null)
                {
                    yield return item;
                }
            }
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        public override string Name
        {
            get { return "Square"; }
        }

        public override FrameworkElement CreateIcon()
        {
            double a = 0.2, b = 0.8;
            return IconBuilder.BuildIcon()
                .Polygon(
                    nameof(AppTheme.ShapeIconFill),
                    nameof(AppTheme.ShapeOutline),
                    new Point(a, a),
                    new Point(b, a),
                    new Point(b, b),
                    new Point(a, b))
                .DependentPoint(a, a)
                .DependentPoint(b, a)
                .Point(b, b)
                .Point(a, b)
                .Canvas;
        }

        public override string HintText
        {
            get
            {
                return "Create a square given two adjacent vertices";
            }
        }
    }
}