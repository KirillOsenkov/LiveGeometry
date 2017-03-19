using System.Windows;
using System.Windows.Controls;
using System.Xml;
using System.Xml.Linq;
using System.Windows.Media;
namespace DynamicGeometry
{
    public class Vector : CompositeFigure, ILengthProvider
    {
        public Vector()
        {
            Line = new Segment();   // Line's dependencies established by OnDependenciesChanged()
            Line.Style = new LineStyle() { StrokeWidth = 0, Color = Colors.Transparent };   // Line is invisible
            Arrow = new Arrow();
            Arrow.ZIndex = (int)ZOrder.Vectors;
            Arrow.Dependencies.Add(Line);
            Children.Add(Line, Arrow);
            ZIndex = (int)ZOrder.Vectors;
        }
        
        public override IFigureStyle Style  // The Arrow's style is used for the vector.
        {
            get
            {
                return Arrow.Style;
            }
            set
            {
                Arrow.Style = value;
            }
        }

        protected override void OnDependenciesChanged()
        {
            base.OnDependenciesChanged();
            Line.Dependencies = Dependencies;
        }

        public override void OnAddingToCanvas(Canvas newContainer)
        {
            base.OnAddingToCanvas(newContainer);
            Arrow.EnsureStyleAssigned();
        }

        public override IFigure HitTest(Point point, System.Predicate<IFigure> filter)
        {
            var result = Arrow.HitTest(point);
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
        }

#endif
        public override void ReadXml(XElement element)
        {
            // Do not use CompositeFigure.ReadXml() because there are no children to read. Children are created by constructor.
            Visible = element.ReadBool("Visible", true);
            Locked = element.ReadBool("Locked", false);
            IsHitTestVisible = element.ReadBool("IsHitTestVisible", true);
            var styleAttribute = element.Attribute("Style");
            if (styleAttribute != null
                && Drawing != null
                && Drawing.StyleManager != null)
            {
                var style = Drawing.StyleManager[styleAttribute.Value];
                if (style != null)
                {
                    this.Arrow.Style = style;
                }
            }
        }

        public PointPair Coordinates
        {
            get { return Line.Coordinates; }
        }

        public double Magnitude
        {
            get { return Line.Length; }
        }

        [PropertyGridVisible]
        [PropertyGridName("Magnitude")]
        public double Length
        {
            get 
            { 
                return Magnitude; 
            }
            set
            {
                Line.Length = value;
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Direction")]
        public double Angle
        {
            get
            {
                return Direction.ToDegrees();
            }
            set
            {
                Line.Angle = value;
            }
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
