using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;

namespace DynamicGeometry
{
    /// <summary>
    /// A segment with an arrowhead. An <see cref="ILine"/> like a segment, so that the tools
    /// that take a line take it (parallel, perpendicular, intersection, a point on it).
    /// </summary>
    public class Vector : CompositeFigure, ILengthProvider, IFixableLength, IFixableDirection, ILine
    {
        public Vector()
        {
            Arrow = new Arrow() { DrawsShaft = false };
            Arrow.Layer = ZOrder.Vectors;
            Line = new VectorShaft(Arrow);   // Line's dependencies established by OnDependenciesChanged()
            Line.Layer = ZOrder.Vectors;
            Arrow.Dependencies.Add(Line);
            Children.Add(Line, Arrow);
            Layer = ZOrder.Vectors;
        }

        /// <summary>
        /// The segment from the start to the end, which tools take for the vector, drawn from
        /// the start to the head: a line in the vector's style, which can be dashed, as the
        /// arrow (a filled outline) could not.
        /// </summary>
        public class VectorShaft : Segment
        {
            readonly Arrow arrow;

            public VectorShaft(Arrow arrow)
            {
                this.arrow = arrow;
            }

            public override void UpdateVisual()
            {
                if (!IsShown || Drawing == null)
                {
                    Shape.Visibility = Visibility.Collapsed;
                    return;
                }

                var outline = arrow.Measure();
                Shape.Set(new PointPair(outline.Tail, outline.HeadBase));
                Shape.Visibility = Visibility.Visible;
            }

            public override void ApplyStyle()
            {
                base.ApplyStyle();

                // the head's color, also for a style whose line is transparent (a vector from
                // an older drawing, in a polygon's style: its fill)
                if (Style != null)
                {
                    Shape.Stroke = DynamicGeometry.Arrow.GetBrush(Style);
                }
            }
        }

        /// <summary>The arrow's style, which the shaft is drawn in too</summary>
        public override IFigureStyle Style
        {
            get
            {
                return Arrow.Style;
            }
            set
            {
                Arrow.Style = value;
                Line.Style = value;
            }
        }

        protected override void OnDependenciesChanged()
        {
            base.OnDependenciesChanged();
            Line.Dependencies = Dependencies;
        }

        // from A to B: vector BA points the other way
        protected override IReadOnlyList<string> NamesFromDependencies()
        {
            return NamesFromPoints(PointOrder.Fixed);
        }

        protected override string Kind
        {
            get
            {
                return "Vector";
            }
        }

        /// <summary>Vector u, v, w... when no points name it</summary>
        protected override string FirstLetter
        {
            get
            {
                return "u";
            }
        }

        public override void OnAddingToCanvas(Canvas newContainer)
        {
            // The arrow is a polygon, and left to itself (which is what the base call does to a
            // child without a style) it would get the default polygon style: a pale translucent
            // fill without an outline. A vector is a line.
            if (Arrow.Style == null && Drawing != null)
            {
                Style = Drawing.StyleManager
                    .GetStyles<LineStyle>()
                    .FirstOrDefault(s => s.GetType() == typeof(LineStyle));
            }

            base.OnAddingToCanvas(newContainer);
            Arrow.EnsureStyleAssigned();
            Line.Style = Arrow.Style;
        }

        // the arrow is a filled polygon, so a hit on it has to land on the drawn pixels: the
        // shaft gets the same room around it as a segment (the invisible line inside)
        public override IFigure HitTest(Point point, System.Predicate<IFigure> filter)
        {
            var result = Arrow.HitTest(point) ?? Line.HitTest(point);
            if (result != null)
            {
                result = this;
                if (!filter(result))
                {
                    result = null;
                }
            }
            return result;
        }

        /// <summary>
        /// Where the vector is, shown or not, as a segment answers: whether a figure is
        /// in view is for whoever asks to decide (<see cref="FigureList.HitTest(Point)"/>
        /// does). A point on a vector and an intersection with one exist where this says
        /// the vector is: asked through the composite's own test, which leaves out what
        /// is hidden, they all stopped existing when the vector was hidden.
        /// </summary>
        public override IFigure HitTest(Point point)
        {
            return HitTest(point, filter: figure => true);
        }

        public Segment Line { get; set; }
        public Arrow Arrow { get; set; }

#if !PLAYER

        public override void WriteXml(XmlWriter writer)
        {
            if (!Visible)
            {
                writer.WriteAttributeString("Visible", "false");
            }
            if (Locked)
            {
                writer.WriteAttributeString("Locked", "true");
            }
            if (Arrow.Style != null)
            {
                writer.WriteAttributeString("Style", Arrow.Style.Name);
            }

            // brought to front or sent to back
            if (Z != 0)
            {
                writer.WriteAttributeString("Z", Z.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }

#endif
        public override void ReadXml(XElement element)
        {
            // Do not use CompositeFigure.ReadXml() because there are no children to read. Children are created by constructor.
            Visible = element.ReadBool("Visible", true);
            Locked = element.ReadBool("Locked", false);
            IsHitTestVisible = element.ReadBool("IsHitTestVisible", true);
            Z = element.ReadInt("Z", 0);
            var styleAttribute = element.Attribute("Style");
            if (styleAttribute != null
                && Drawing != null
                && Drawing.StyleManager != null)
            {
                var style = Drawing.StyleManager[styleAttribute.Value];
                if (style != null)
                {
                    Style = style;
                }
            }
        }

        public PointPair Coordinates
        {
            get { return Line.Coordinates; }
        }

        // the length lives on the segment inside; so do Fix length and Free length
        [PropertyGridVisible]
        [PropertyGridGroup("Length")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(LengthPropertyValue))]
        public double Length
        {
            get
            {
                return Line.Length;
            }
            set
            {
                Line.Length = value;
            }
        }

        // the direction too: an end that can't be turned (a point on a figure, a start built on
        // the end) leaves the row read-only, where a set was an undo step that undid nothing
        public bool CanEdit(string propertyName)
        {
            return Line.CanEdit(propertyName);
        }

        public IPoint LengthEndpoint => Line.LengthEndpoint;

        public IPoint DirectionEndpoint => Line.DirectionEndpoint;

        public IList<IFigure> MeasuredFigures
        {
            get { return new IFigure[] { this }; }
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

        [PropertyGridVisible]
        [PropertyGridName("Fix length")]
        [PropertyGridGroup("Length")]
        [PropertyGridIcon(PropertyGridIcon.Lock)]
        public void FixLength()
        {
            Line.FixLength();
            Drawing.RaiseDisplayProperties(this);
        }

        [PropertyGridVisible]
        [PropertyGridName("Free length")]
        [PropertyGridGroup("Length")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void FreeLength()
        {
            Line.FreeLength();
            Drawing.RaiseDisplayProperties(this);
        }

        // undo puts the end back where it was: turned back by the angle, it came back to its
        // place only while the start had stayed put
        [PropertyGridVisible]
        [PropertyGridName("Direction")]
        [PropertyGridGroup("Direction")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(DirectionPropertyValue))]
        public double Angle
        {
            get
            {
                return Line.Angle;
            }
            set
            {
                Line.Angle = value;
            }
        }

        // so do Fix direction, Free direction, Horizontal and Vertical
        [PropertyGridVisible]
        [PropertyGridName("Fix direction")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Lock)]
        public void FixDirection()
        {
            Line.FixDirection();
            Drawing.RaiseDisplayProperties(this);
        }

        [PropertyGridVisible]
        [PropertyGridName("Free direction")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void FreeDirection()
        {
            Line.FreeDirection();
            Drawing.RaiseDisplayProperties(this);
        }

        [PropertyGridVisible]
        [PropertyGridName("Horizontal")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Horizontal)]
        public void MakeHorizontal()
        {
            Line.MakeHorizontal();
            Drawing.RaiseDisplayProperties(this);
        }

        [PropertyGridVisible]
        [PropertyGridName("Vertical")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Vertical)]
        public void MakeVertical()
        {
            Line.MakeVertical();
            Drawing.RaiseDisplayProperties(this);
        }

        public double Direction
        {
            get 
            { 
                return Math.GetAngle(Coordinates.P1, Coordinates.P2); 
            }
        }

        public double GetNearestParameterFromPoint(Point point)
        {
            return Line.GetNearestParameterFromPoint(point);
        }

        public Point GetPointFromParameter(double parameter)
        {
            return Line.GetPointFromParameter(parameter);
        }

        public Tuple<double, double> GetParameterDomain()
        {
            return Line.GetParameterDomain();
        }

        public override string ToString()
        {
            // See comment in Segment.ToString() - D.H.
            return Name;
            //return "Vector " + Dependencies[0].ToString() + Dependencies[1].ToString();
        }
        
    }
}
