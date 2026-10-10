using Avalonia.Media;
using Avalonia.Controls.Shapes;

namespace DynamicGeometry
{

    public partial class EllipseArc : EllipseArcBase
    {
        protected override string Kind
        {
            get
            {
                return "Elliptical arc";
            }
        }

        public override int BeginPointIndex
        {
            get { return 3; }
        }

        public override int EndPointIndex
        {
            get { return 4; }
        }

        /// <summary>
        /// A Convert verb of the figure's own grid: the new figure's grid shows after it, where
        /// the old one's was (the verbs of <see cref="ArcPanel"/> go through <see cref="Replace"/>
        /// and show the panel again)
        /// </summary>
        public static void Convert(IArc oldArc, IArc newArc)
        {
            Replace(oldArc, newArc);
            newArc.Drawing.RaiseDisplayProperties(newArc);
        }

        /// <summary>
        /// The new arc, sector or segment in the old one's place, one undo step: same points,
        /// same direction, hidden or locked if the old one was, and the style traded for the
        /// filled or unfilled one of its hue (<see cref="StyleManager.ConvertStyle"/>): a sector
        /// made of an arc is seen at once, an arc made of a sector is a line again
        /// </summary>
        public static void Replace(IArc oldArc, IArc newArc)
        {
            var drawing = oldArc.Drawing;
            newArc.Style = drawing.StyleManager.ConvertStyle(oldArc, newArc);
            newArc.Clockwise = oldArc.Clockwise;

            // a hidden helper (converted from the Figure List) stays hidden
            newArc.Visible = oldArc.Visible;
            newArc.Locked = oldArc.Locked;
            Actions.ReplaceWithNew(oldArc, newArc);
            drawing.RaiseUserIsAddingFigures(new Drawing.UIAFEventArgs() { Figures = newArc.AsEnumerable<IFigure>() });
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to segment")]
        [PropertyGridIcon(PropertyGridIcon.CircleSegment)]
        public virtual void ConvertToEllipseSegment()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseSegment(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to sector")]
        [PropertyGridIcon(PropertyGridIcon.Sector)]
        public void ConvertToEllipseSector()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseSector(this.Drawing, this.Dependencies));
        }

#endif

    }

    public partial class CircleArc : CircleArcBase
    {
        protected override string Kind
        {
            get
            {
                return "Arc";
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to segment")]
        [PropertyGridIcon(PropertyGridIcon.CircleSegment)]
        public virtual void ConvertToCircleSegment()
        {
            EllipseArc.Convert(this, Factory.CreateCircleSegment(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to sector")]
        [PropertyGridIcon(PropertyGridIcon.Sector)]
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
    public partial class CircleSegment : CircleArcBase, IShapeWithInterior, IPerimeter, IConditionalProperties
    {
        protected override string Kind
        {
            get
            {
                return "Circular segment";
            }
        }

        /// <summary>No "Convert to arc" while something measures the area or the perimeter: a bare arc has neither</summary>
        public bool CanEdit(string propertyName)
        {
            return propertyName != "ConvertToArc" || !this.IsUsedForArea();
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }


        protected override Path CreateShape()
        {
            var result = base.CreateShape();
            Figure.IsClosed = true;
            return result;
        }

        // (it said r² * angle / π: a half disc came out as r², not π r² / 2)
        public double Area
        {
            get
            {
                return SegmentArea;
            }
        }

        /// <summary>The arc and its chord</summary>
        public double Perimeter
        {
            get
            {
                return Length + BeginLocation.Distance(EndLocation);
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to arc")]
        [PropertyGridIcon(PropertyGridIcon.Arc)]
        public void ConvertToArc()
        {
            EllipseArc.Convert(this, Factory.CreateArc(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to sector")]
        [PropertyGridIcon(PropertyGridIcon.Sector)]
        public void ConvertToSector()
        {
            EllipseArc.Convert(this, Factory.CreateCircleSector(this.Drawing, this.Dependencies));
        }

#endif

    }

    public partial class EllipseSegment : EllipseArcBase, IShapeWithInterior, IPerimeter, IConditionalProperties
    {
        protected override string Kind
        {
            get
            {
                return "Elliptical segment";
            }
        }

        /// <summary>The arc and its chord</summary>
        public double Perimeter
        {
            get
            {
                return Length + BeginLocation.Distance(EndLocation);
            }
        }

        /// <summary>No "Convert to arc" while something measures the area or the perimeter: a bare arc has neither</summary>
        public bool CanEdit(string propertyName)
        {
            return propertyName != "ConvertToEllipseArc" || !this.IsUsedForArea();
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

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

        // (it was "not a number": the Area tool on an elliptical segment said NaN)
        public double Area
        {
            get
            {
                return SegmentArea;
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to arc")]
        [PropertyGridIcon(PropertyGridIcon.Arc)]
        public void ConvertToEllipseArc()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseArc(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to sector")]
        [PropertyGridIcon(PropertyGridIcon.Sector)]
        public void ConvertToEllipseSector()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseSector(this.Drawing, this.Dependencies));
        }

#endif

    }

    public partial class CircleSector : CircleArcBase, IShapeWithInterior, IPerimeter, IConditionalProperties
    {
        protected override string Kind
        {
            get
            {
                return "Sector";
            }
        }

        /// <summary>The arc and the two radii</summary>
        public double Perimeter
        {
            get
            {
                return Length + Center.Distance(BeginLocation) + Center.Distance(EndLocation);
            }
        }

        /// <summary>No "Convert to arc" while something measures the area or the perimeter: a bare arc has neither</summary>
        public bool CanEdit(string propertyName)
        {
            return propertyName != "ConvertToArc" || !this.IsUsedForArea();
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

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

        // (it said r² * angle / π plus the triangle: a quarter disc came out as r², not π r² / 4)
        public double Area
        {
            get
            {
                return SectorArea;
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to arc")]
        [PropertyGridIcon(PropertyGridIcon.Arc)]
        public void ConvertToArc()
        {
            EllipseArc.Convert(this, Factory.CreateArc(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to segment")]
        [PropertyGridIcon(PropertyGridIcon.CircleSegment)]
        public virtual void ConvertToCircleSegment()
        {
            EllipseArc.Convert(this, Factory.CreateCircleSegment(this.Drawing, this.Dependencies));
        }

#endif

    }

    public partial class EllipseSector : EllipseArcBase, IShapeWithInterior, IPerimeter, IConditionalProperties
    {
        protected override string Kind
        {
            get
            {
                return "Elliptical sector";
            }
        }

        /// <summary>The arc and the two radii</summary>
        public double Perimeter
        {
            get
            {
                return Length + Center.Distance(BeginLocation) + Center.Distance(EndLocation);
            }
        }

        /// <summary>No "Convert to arc" while something measures the area or the perimeter: a bare arc has neither</summary>
        public bool CanEdit(string propertyName)
        {
            return propertyName != "ConvertToEllipseArc" || !this.IsUsedForArea();
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

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
                return SectorArea;
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to arc")]
        [PropertyGridIcon(PropertyGridIcon.Arc)]
        public void ConvertToEllipseArc()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseArc(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to segment")]
        [PropertyGridIcon(PropertyGridIcon.CircleSegment)]
        public virtual void ConvertToEllipseSegment()
        {
            EllipseArc.Convert(this, Factory.CreateEllipseSegment(this.Drawing, this.Dependencies));
        }

#endif

    }

}
