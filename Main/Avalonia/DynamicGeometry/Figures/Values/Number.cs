using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;

namespace DynamicGeometry
{
    /// <summary>
    /// A number in the drawing: a figure without a shape that other figures depend on the way
    /// they depend on a segment's length. It is what a typed value becomes (the distance of a
    /// translation), so that it can be edited later, shared by several figures and named in an
    /// expression. Provides itself as a length (in units) and as an angle (in degrees). One
    /// created on demand for a single figure is <see cref="FigureBase.Auxiliary"/> and leaves
    /// the drawing with its last user.
    /// </summary>
    public class Number : FigureBase, INumber, ILengthProvider, IAngleProvider, IConditionalProperties
    {
        /// <summary>A number has no shape and so no style: the style buttons every figure has are not its (they threw)</summary>
        public bool CanEdit(string propertyName)
        {
            return propertyName != nameof(EditStyleButton) && propertyName != nameof(CreateNewStyle);
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

        public Number()
        {
            IsHitTestVisible = false;
        }

        // Nothing on screen to show, hide, lock or style: no rows for them (they stood empty
        // or did nothing).

        [PropertyGridVisible(false)]
        public override IFigureStyle StyleDisplay
        {
            get
            {
                return base.StyleDisplay;
            }
            set
            {
                base.StyleDisplay = value;
            }
        }

        [PropertyGridVisible(false)]
        public override bool Visible
        {
            get
            {
                return base.Visible;
            }
            set
            {
                base.Visible = value;
            }
        }

        [PropertyGridVisible(false)]
        public override bool Locked
        {
            get
            {
                return base.Locked;
            }
            set
            {
                base.Locked = value;
            }
        }

        public static Number CreateAuxiliary(Drawing drawing, double value)
        {
            return new Number() { Drawing = drawing, Auxiliary = true, Value = value };
        }

        double value;

        [PropertyGridVisible]
        [PropertyGridPreferredEditor("UpDown")]
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

        protected override string Kind
        {
            get
            {
                return "Number";
            }
        }

        /// <summary>A number goes by its name: "radius n1"</summary>
        public override string Noun
        {
            get
            {
                return null;
            }
        }

        /// <summary>"= 3"</summary>
        public override string Construction
        {
            get
            {
                return "= " + ConstructionText.Number(Value);
            }
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
