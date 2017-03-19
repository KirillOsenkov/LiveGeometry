using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;

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

            var side0 = Factory.CreateSegment(Drawing, p1, p2);
            var circle = Factory.CreateCircle(Drawing, new[] { p2, p1 });
            var perpendicular = Factory.CreatePerpendicularLine(Drawing, new IFigure[] { side0, p2 });
            var intersection = Factory.CreateIntersectionPoint(Drawing, circle, perpendicular,
                perpendicular.Coordinates.P2);
            var midpoint = Factory.CreateMidPoint(Drawing, new IFigure[] { intersection, p1 });
            var reflectedPoint = Factory.CreateReflectedPoint(Drawing, new IFigure[] { p2, midpoint });
            var side1 = Factory.CreateSegment(Drawing, p2, intersection);
            var side2 = Factory.CreateSegment(Drawing, intersection, reflectedPoint);
            var side3 = Factory.CreateSegment(Drawing, reflectedPoint, p1);
            var polygon = Factory.CreatePolygon(Drawing, new IFigure[] { p1, p2, intersection, reflectedPoint });
            var midpoint0 = Factory.CreateMidPoint(Drawing, side0);
            var midpoint1 = Factory.CreateMidPoint(Drawing, side1);
            var midpoint2 = Factory.CreateMidPoint(Drawing, side2);
            var midpoint3 = Factory.CreateMidPoint(Drawing, side3);

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
                yield return item;
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
                    new SolidColorBrush(Color.FromArgb(255, 128, 255, 128)),
                    new SolidColorBrush(Colors.Black),
                    new Point(a, a),
                    new Point(b, a),
                    new Point(b, b),
                    new Point(a, b))
                .Line(a, a, b, a)
                .Line(b, a, b, b)
                .Line(b, b, a, b)
                .Line(a, b, a, a)
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