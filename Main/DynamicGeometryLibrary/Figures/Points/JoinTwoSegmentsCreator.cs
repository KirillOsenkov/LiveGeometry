using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Lines)]
    [Order(11)]
    public class JoinTwoSegmentsCreator : Behavior
    {
        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            var coordinates = Coordinates(e);
            var underMouse = Drawing.Figures.HitTest<FreePoint>(coordinates);
            if (underMouse != null)
            {
                JoinSegments(underMouse);
                // for polyline
                JoinPolyLineSegments(underMouse);
                RemovePointFromPolygons(underMouse);
            }
        }

        void JoinSegments(FreePoint point)
        {
            var dependents = point.Dependents.OfType<Segment>().ToArray();
            if (dependents.Length != 2)
            {
                return;
            }
            var line1 = dependents[0];
            var line2 = dependents[1];

            var otherPoint1 = line1.Dependencies.Without(point).FirstOrDefault();
            var otherPoint2 = line2.Dependencies.Without(point).FirstOrDefault();
            if (otherPoint1 == null || otherPoint2 == null)
            {
                return;
            }

            var segment = Factory.CreateSegment(Drawing, new[] { otherPoint1, otherPoint2 });

            using (Drawing.ActionManager.CreateTransaction())
            {
                RemovePointFromPolygons(point);
                Actions.Remove(line2);
                Actions.ReplaceWithNew(line1, segment);
                Actions.Remove(point);
            }
        }

        void RemovePointFromPolygons(FreePoint point)
        {
            foreach (var polygon in point.Dependents.OfType<PolygonBase>().ToList())
            {
                if (polygon.Dependencies.Count > 3)
                {
                    RemovePointFromPolygon(point, polygon);
                }
            }
        }

        void RemovePointFromPolygon(FreePoint point, PolygonBase polygon)
        {
            Actions.RemoveDependency(polygon, point);
        }

        public override string Name
        {
            get { return "Join segments"; }
        }

        public override string HintText
        {
            get
            {
                return "Click a point between two segments to join the other two points.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Line(0.25, 0.75, 0.4, 0.4)
                .Line(0.4, 0.4, 0.75, 0.25)
                .Point(0.25, 0.75)
                .Point(0.75, 0.25)
                .Canvas;
        }

        void JoinPolyLineSegments(FreePoint point)
        {
            var dependents = point.Dependents.OfType<Polyline>().ToArray();

            foreach (object obj in dependents)
            {
                if (obj is Polyline)
                {
                    Polyline polyline = (Polyline)obj;

                    if (polyline.Dependencies.Count <= 3)
                    {
                        return;
                    }

                    List<IFigure> NewPolyLinePoints = new List<IFigure>();

                    // Eliminate deleted point
                    for (int i = 0; i < polyline.Dependencies.Count; i++)
                    {
                        IPoint p1 = polyline.Dependencies[i] as IPoint;

                        if (p1.Coordinates.X != point.Coordinates.X
                            && p1.Coordinates.Y != point.Coordinates.Y)
                        {
                            NewPolyLinePoints.Add(p1);
                        }
                    }

                    // create new polyline
                    var newPolyLine = Factory.CreatePolyline(Drawing, NewPolyLinePoints);
                    using (Drawing.ActionManager.CreateTransaction())
                    {
                        Actions.Remove(point);
                        Actions.ReplaceWithNew(polyline, newPolyLine);
                    }
                }
            }
        }
    }
}