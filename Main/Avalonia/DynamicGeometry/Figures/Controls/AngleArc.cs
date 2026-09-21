using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace DynamicGeometry
{
    public class AngleArc : CircleArc
    {
        public AngleArc()
            : base()
        {
            Size = 16;
            ArcShape.Size = new Size(this.Size, this.Size);
        }

        public const double DefaultSize = 16;

        double size = DefaultSize;

        /// <summary>Radius of the (first) arc in pixels</summary>
        [PropertyGridVisible]
        [PropertyGridName("Radius")]
        [Domain(10, 100)]
        public double Size
        {
            get
            {
                return size;
            }
            set
            {
                size = value;
                UpdateIfInDrawing();
            }
        }

        int arcCount = 1;

        /// <summary>
        /// How the angle is marked: not at all, ), )) or ))) - to tell equal angles from the
        /// others, like the original DG could.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Arcs")]
        [Domain(0, 3)]
        public int ArcCount
        {
            get
            {
                return arcCount;
            }
            set
            {
                arcCount = System.Math.Max(0, System.Math.Min(3, value));
                UpdateIfInDrawing();
            }
        }

        /// <summary>Gap between two arcs in pixels, on top of the stroke width</summary>
        public const double ArcSpacing = 2;

        void UpdateIfInDrawing()
        {
            if (Drawing != null && Shape != null)
            {
                UpdateVisual();
            }
        }

        class ExtraArc
        {
            public Avalonia.Media.PathFigure Figure;
            public Avalonia.Media.ArcSegment Segment;
        }

        readonly List<ExtraArc> extraArcs = new List<ExtraArc>();
        int shownFigureCount = 1;

        /// <summary>
        /// The path has one figure per arc: the one the base class made, plus these.
        /// </summary>
        void UpdateFigureCount(int count)
        {
            if (count == shownFigureCount)
            {
                return;
            }

            shownFigureCount = count;
            while (extraArcs.Count < count - 1)
            {
                var segment = new Avalonia.Media.ArcSegment();
                var figure = new Avalonia.Media.PathFigure()
                {
                    IsClosed = false,
                    IsFilled = false,
                    Segments = new Avalonia.Media.PathSegments() { segment }
                };
                extraArcs.Add(new ExtraArc() { Figure = figure, Segment = segment });
            }

            while (extraArcs.Count > System.Math.Max(0, count - 1))
            {
                extraArcs.RemoveAt(extraArcs.Count - 1);
            }

            // the first one always stays (without segments when there is nothing to draw)
            var figures = new Avalonia.Media.PathFigures();
            figures.Add(Figure);
            foreach (var arc in extraArcs)
            {
                figures.Add(arc.Figure);
            }

            ((Avalonia.Media.PathGeometry)Shape.Data).Figures = figures;
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            var arcs = element.Attribute("Arcs");
            int count;
            if (arcs != null && int.TryParse(arcs.Value, out count))
            {
                arcCount = System.Math.Max(0, System.Math.Min(3, count));
            }

            if (element.Attribute("Radius") != null)
            {
                size = System.Math.Max(10, System.Math.Min(100, element.ReadDouble("Radius")));
            }
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            if (ArcCount != 1)
            {
                writer.WriteAttributeString("Arcs", ArcCount.ToString());
            }

            if (Size != DefaultSize)
            {
                writer.WriteAttributeDouble("Radius", Size);
            }
        }

        /// <summary>
        /// In radians: 0.005 degrees, which is when the label (two decimals) reads "90°".
        /// The sign and the number never disagree; an angle that merely looks right doesn't get it.
        /// </summary>
        public const double RightAngleTolerance = 0.005 * System.Math.PI / 180;

        readonly Avalonia.Media.LineSegment rightAngleSide1 = new Avalonia.Media.LineSegment();
        readonly Avalonia.Media.LineSegment rightAngleSide2 = new Avalonia.Media.LineSegment();

        enum Sign
        {
            None,
            Arc,
            RightAngle
        }

        Sign shownSign = Sign.Arc;

        public override double Radius
        {
            get
            {
                return ToLogical(Size);
            }
        }

        public override double SemiMajor
        {
            get
            {
                return Radius;
            }
        }

        public override double SemiMinor
        {
            get
            {
                return Radius;
            }
        }

        public override Point BeginLocation
        {
            get
            {
                return Math.ScalePointBetweenTwo(
                Center,
                Point(1),
                Radius / Center.Distance(Point(1)));
            }
        }
        
        [PropertyGridVisible]
        public virtual double Measure
        {
            get 
            {
                return Angle;
                //return Math.OAngle(BeginLocation, Center, EndLocation); 
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to opposite angle")]
        public void ConvertToOpposite()
        {
            ConvertToOpposite(this);
        }

        /// <summary>
        /// An angle is two figures, the arc and the number, on the same vertex and the same two
        /// points. This finds the other one of the pair (null if it was deleted).
        /// </summary>
        public static IFigure FindCompanion(IFigure angleFigure)
        {
            var dependencies = angleFigure.Dependencies;
            if (dependencies.Count != 3)
            {
                return null;
            }

            bool wantArc = angleFigure is AngleMeasurement;
            return dependencies[0].Dependents.FirstOrDefault(f =>
                f != angleFigure
                && (wantArc ? f is AngleArc : f is AngleMeasurement)
                && f.Dependencies.Count == 3
                && f.Dependencies[0] == dependencies[0]
                && f.Dependencies.Contains(dependencies[1])
                && f.Dependencies.Contains(dependencies[2]));
        }

        /// <summary>
        /// Swaps the two sides, which turns the angle into the one that completes it to 360
        /// degrees - for the arc and the number together, whichever of them was asked.
        /// </summary>
        public static void ConvertToOpposite(IFigure angleFigure)
        {
            var dependencies = angleFigure.Dependencies;
            if (dependencies.Count != 3)
            {
                return;
            }

            var companion = FindCompanion(angleFigure);

            var side = dependencies[1];
            dependencies[1] = dependencies[2];
            dependencies[2] = side;
            angleFigure.RecalculateAndUpdateVisual();

            // not "swap it too": make it the same, in case the two were out of step already
            // (an array says IsReadOnly, but its elements can be set, and arrays are what these are)
            if (companion != null)
            {
                companion.Dependencies[1] = dependencies[1];
                companion.Dependencies[2] = dependencies[2];
                companion.RecalculateAndUpdateVisual();
            }
        }

        /// <summary>
        /// Anywhere on the mark: from the first arc out to the last one, the gaps between them
        /// included - not just the first arc, which is all the base class knows about.
        /// </summary>
        public override IFigure HitTest(Point point)
        {
            var found = base.HitTest(point);
            if (found != null || ArcCount < 2 || shownSign != Sign.Arc)
            {
                return found;
            }

            var center = Center;
            var distance = center.Distance(point);
            var tolerance = CursorTolerance + LogicalWidth() / 2;
            var outerRadius = ToLogical(Size + (ArcCount - 1) * (Shape.StrokeThickness + ArcSpacing));
            if (distance < Radius - tolerance || distance > outerRadius + tolerance)
            {
                return null;
            }

            var angleToPoint = Math.GetAngle(center, point);
            return Math.IsAngleBetweenAngles(angleToPoint, StartAngle, EndAngle, Clockwise) ? this : null;
        }

        public override void UpdateVisual()
        {
            var center = Point(0);
            var distance1 = center.Distance(Point(1));
            var distance2 = center.Distance(Point(2));
            if (distance1 == 0 || distance2 == 0)
            {
                Shape.Visibility = Visibility.Collapsed;
                return;
            }

            var angle = Math.OAngle(BeginLocation, center, EndLocation);
            var isRightAngle = System.Math.Abs(angle - Math.PI / 2) < RightAngleTolerance;
            // What the first figure of the path is made of. "Nothing" is a figure without
            // segments and not a path without figures: an empty path doesn't get repainted.
            var sign = ArcCount == 0 ? Sign.None : (isRightAngle ? Sign.RightAngle : Sign.Arc);
            if (sign != shownSign)
            {
                shownSign = sign;
                var segments = new Avalonia.Media.PathSegments();
                if (sign == Sign.RightAngle)
                {
                    segments.Add(rightAngleSide1);
                    segments.Add(rightAngleSide2);
                }
                else if (sign == Sign.Arc)
                {
                    segments.Add(ArcShape);
                }

                Figure.Segments = segments;
            }

            // a right angle has one sign however many arcs were asked for, like in the original DG
            int figuresNeeded = isRightAngle ? System.Math.Min(ArcCount, 1) : ArcCount;
            UpdateFigureCount(figuresNeeded);

            var corner = ToPhysical(center);
            var first = RightAngleMark.Direction(corner, ToPhysical(Point(1)));
            var second = RightAngleMark.Direction(corner, ToPhysical(Point(2)));
            if (first == null || second == null)
            {
                return;
            }

            if (isRightAngle)
            {
                // the school sign for 90 degrees: two sides of a little square, not an arc
                var points = RightAngleMark.GetPoints(corner, first.Value, second.Value, RightAngleMark.Size);
                Figure.StartPoint = points[0];
                rightAngleSide1.Point = points[1];
                rightAngleSide2.Point = points[2];
                return;
            }

            ArcShape.Size = new Size(Size, Size);
            Figure.StartPoint = ToPhysical(BeginLocation);
            ArcShape.Point = ToPhysical(EndLocation);
            ArcShape.IsLargeArc = angle > Math.PI;

            // the second and third arc, each a stroke and a bit further out
            for (int i = 0; i < extraArcs.Count; i++)
            {
                var radius = Size + (i + 1) * (Shape.StrokeThickness + ArcSpacing);
                var arc = extraArcs[i];
                arc.Figure.StartPoint = corner + first.Value * radius;
                arc.Segment.Point = corner + second.Value * radius;
                arc.Segment.Size = new Size(radius, radius);
                arc.Segment.IsLargeArc = ArcShape.IsLargeArc;
                arc.Segment.SweepDirection = ArcShape.SweepDirection;
            }

            // I commented out these lines because this causes AngleArc not to honor Visible property. Is this a mistake?  
            // I believe there will be times when a user would like to hide the arc. D. H.
            //if (Shape.Visibility != Visibility.Visible)
            //{
            //    Shape.Visibility = Visibility.Visible;
            //}
        }
    }
}
