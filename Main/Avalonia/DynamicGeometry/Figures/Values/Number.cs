using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;

namespace DynamicGeometry
{
    /// <summary>
    /// A number in the drawing: a figure without a shape that other figures depend on the way
    /// they depend on a segment's length. It is what a typed value becomes (the magnitude of a
    /// translation), so that it can be edited later, shared by several figures and named in an
    /// expression. Provides itself as a length (in units) and as an angle (in degrees). One
    /// created on demand for a single figure is <see cref="FigureBase.Auxiliary"/> and leaves
    /// the drawing with its last user.
    /// </summary>
    public class Number : FigureBase, INumber, ILengthProvider, IAngleProvider
    {
        public Number()
        {
            IsHitTestVisible = false;
        }

        public static Number CreateAuxiliary(Drawing drawing, double value)
        {
            return new Number() { Drawing = drawing, Auxiliary = true, Value = value };
        }

        double value;

        [PropertyGridVisible]
        public double Value
        {
            get
            {
                return value;
            }
            set
            {
                this.value = value;
                RaisePropertyChanged("Value");
                if (Drawing != null && !Dependents.IsEmpty())
                {
                    this.RecalculateAllDependents();
                }
            }
        }

        public double Length
        {
            get { return Value; }
        }

        /// <summary>Angle providers speak radians; the number itself is in degrees</summary>
        public double Angle
        {
            get { return Value.ToRadians(); }
        }

        // n1, n2, n3: short, since they are what an expression or a "tied to" row shows
        public override string GenerateFigureName(List<string> blacklist)
        {
            for (int i = 1; ; i++)
            {
                var candidate = "n" + i;
                if (this.NameAvailable(candidate) && (blacklist == null || !blacklist.Contains(candidate)))
                {
                    return candidate;
                }
            }
        }

        public override void ApplyStyle()
        {
        }

        public override IFigure HitTest(Avalonia.Point point)
        {
            return null;
        }

        public override void ReadXml(XElement element)
        {
            base.ReadXml(element);
            Value = element.ReadDouble("Value");
        }

        public override void WriteXml(XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeDouble("Value", Value);
        }
    }
}
