using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Collections.Generic;
using System.Linq;
using M = System.Math;

namespace DynamicGeometry
{

    public partial class EllipseArc : EllipseArcBase
    {

        public override int BeginPointIndex
        {
            get { return 3; }
        }

        public override int EndPointIndex
        {
            get { return 4; }
        }

        public static void Convert(IArc oldArc, IArc newArc)
        {
            var drawing = oldArc.Drawing;
            newArc.Style = oldArc.Style;
            newArc.Clockwise = oldArc.Clockwise;
            Actions.ReplaceWithNew(oldArc, newArc);
            drawing.RaiseUserIsAddingFigures(new Drawing.UIAFEventArgs() { Figures = newArc.AsEnumerable<IFigure>() });
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert To Segment")]
        public virtual void ConvertToEllipseSegment()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseSegment(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert To Sector")]
        public void ConvertToEllipseSector()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseSector(this.Drawing, this.Dependencies));
        }

#endif

    }

    public partial class CircleArc : CircleArcBase
    {

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert To Segment")]
        public virtual void ConvertToCircleSegment()
        {
            EllipseArc.Convert(this, Factory.CreateCircleSegment(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert To Sector")]
        public void ConvertToSector()
        {
            EllipseArc.Convert(this, Factory.CreateCircleSector(this.Drawing, this.Dependencies));
        }

#endif

    }

    // Below are simple implementations of a circle and ellipse segments.  
    // The chord is represented visually but is not a functioning figure in the drawing.
    // A segment could be added to this figure to provide a functioning chord. (SquareCreator is a model to follow.)
    // Implementing this as a composite figure is probably not a good idea. Intersections, pointOnFigure, etc would be ambiguous.
    // Unlike an arc, a circle or ellipse segment has a defined area.
    public partial class CircleSegment : CircleArcBase, IShapeWithInterior
    {
        
        protected override Path CreateShape()
        {
            var result = base.CreateShape();
            Figure.IsClosed = true;
            return result;
        }

        public double Area
        {
            get
            {
                return Radius.Sqr() * Angle / Math.PI;
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert To Arc")]
        public void ConvertToArc()
        {
            EllipseArc.Convert(this, Factory.CreateArc(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert To Sector")]
        public void ConvertToSector()
        {
            EllipseArc.Convert(this, Factory.CreateCircleSector(this.Drawing, this.Dependencies));
        }

#endif

    }

    public partial class EllipseSegment : EllipseArcBase, IShapeWithInterior
    {
        protected override Path CreateShape()
        {
            var result = base.CreateShape();
            Figure.IsClosed = true;
            return result;
        }

        public override int BeginPointIndex
        {
            get { return 3; }
        }

        public override int EndPointIndex
        {
            get { return 4; }
        }

        public double Area
        {
            get
            {
                // Area of ellipse sector. Having trouble with integration.
                return double.NaN;

                //double t1 = StartAngle;
                //double t2 = EndAngle;
                //double a = SemiMajor;
                //double b = SemiMinor;
                //Point C = Center;
                //double angle = Angle;
                //var canonicalEndPoint = Math.RotatePoint(EndLocation, C, -angle).Minus(C);
                //var canonicalBeginPoint = Math.RotatePoint(BeginLocation, C, -angle).Minus(C);
                //t1 = M.Acos(canonicalBeginPoint.X / SemiMajor);
                //if (StartAngle > Math.PI) t1 += Math.PI;
                //t2 = M.Acos(canonicalEndPoint.X / SemiMajor);
                //if (EndAngle > Math.PI) t2 += Math.PI;
                //double sectorArea = 0;
                ////sectorArea = .5 * SemiMajor * SemiMinor * (t2 - t1);
                //if (Clockwise)
                //{
                //    var temp = t1;
                //    t1 = t2;
                //    t2 = temp;
                //}
                //if (t2 > t1)
                //{
                //    sectorArea = (t2 - t1) * (a * a + b * b) / 4 + (M.Sin(t2) * M.Cos(t2) - M.Sin(t1) * M.Cos(t1)) * (a * a - b * b) / 4;
                //}
                //else
                //{
                //    // Can't integrate across a disconinuity. Integrate in parts.
                //    double dp = Math.DOUBLEPI;
                //    double area1 = (dp - t1) * (a * a + b * b) / 4 + (0 - M.Sin(t1) * M.Cos(t1)) * (a * a - b * b) / 4;
                //    double area2 = (t2) * (a * a + b * b) / 4 + (M.Sin(t2) * M.Cos(t2) - 0) * (a * a - b * b) / 4;
                //    sectorArea = area1 + area2;
                //}

                //// Area of Triangle
                //Point B = BeginLocation;
                //Point E = EndLocation;
                //double triangleArea = M.Abs((C.X * B.Y - B.X * C.Y) / 2 + (B.X * E.Y - E.X * B.Y) / 2 + (E.X * C.Y - C.X * E.Y) / 2);

                //// The triangle area should be added to the sector area for large arcs.
                //if (M.Sin(ArcAngle) < 0)
                //{
                //    triangleArea = -triangleArea;
                //}
                //return M.Abs(sectorArea) - triangleArea;
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert To Arc")]
        public void ConvertToEllipseArc()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseArc(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert To Sector")]
        public void ConvertToEllipseSector()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseSector(this.Drawing, this.Dependencies));
        }

#endif

    }

    public partial class CircleSector : CircleArcBase, IShapeWithInterior
    {
        private PathFigure PolygonPart;
        private LineSegment Side1;
        private LineSegment Side2;
        protected override Path CreateShape()
        {
            var result = base.CreateShape();
            Side1 = new LineSegment();
            Side2 = new LineSegment();
            PolygonPart = new PathFigure()
            {
                IsClosed = false,
                IsFilled = true,
                Segments = new PathSegmentCollection()
                {
                    Side1,
                    Side2
                }
            };
            (result.Data as PathGeometry).Figures.Add(PolygonPart);
            result.StrokeEndLineCap = PenLineCap.Round;
            result.StrokeStartLineCap = PenLineCap.Round;
            return result;
        }

        public override void UpdateVisual()
        {
            base.UpdateVisual();
            PolygonPart.StartPoint = ToPhysical(BeginLocation);
            Side1.Point = ToPhysical(Center);
            Side2.Point = ToPhysical(EndLocation);
        }

        public double Area
        {
            get
            {
                var segmentArea = Radius.Sqr() * Angle / Math.PI;
                var polygonArea = Math.Area(BeginLocation, Center, EndLocation);
                return segmentArea + polygonArea;
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert To Arc")]
        public void ConvertToArc()
        {
            EllipseArc.Convert(this, Factory.CreateArc(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert To Segment")]
        public virtual void ConvertToCircleSegment()
        {
            EllipseArc.Convert(this, Factory.CreateCircleSegment(this.Drawing, this.Dependencies));
        }

#endif

    }

    public partial class EllipseSector : EllipseArcBase, IShapeWithInterior
    {
        private PathFigure PolygonPart;
        private LineSegment Side1;
        private LineSegment Side2;
        protected override Path CreateShape()
        {
            var result = base.CreateShape();
            Side1 = new LineSegment();
            Side2 = new LineSegment();
            PolygonPart = new PathFigure()
            {
                IsClosed = false,
                IsFilled = true,
                Segments = new PathSegmentCollection()
                {
                    Side1,
                    Side2
                }
            };
            (result.Data as PathGeometry).Figures.Add(PolygonPart);
            result.StrokeEndLineCap = PenLineCap.Round;
            result.StrokeStartLineCap = PenLineCap.Round;
            return result;
        }

        
        public override void UpdateVisual()
        {
            base.UpdateVisual();
            PolygonPart.StartPoint = ToPhysical(BeginLocation);
            Side1.Point = ToPhysical(Center);
            Side2.Point = ToPhysical(EndLocation);
        }

        public override int BeginPointIndex
        {
            get { return 3; }
        }

        public override int EndPointIndex
        {
            get { return 4; }
        }

        public double Area
        {
            get
            {
                // Area of ellipse sector. Having trouble with integration.
                return double.NaN;
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert To Arc")]
        public void ConvertToEllipseArc()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseArc(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert To Segment")]
        public virtual void ConvertToEllipseSegment()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseSegment(this.Drawing, this.Dependencies));
        }

#endif

    }

}
