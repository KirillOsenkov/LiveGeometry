using System.Linq;
using Avalonia;
using Avalonia.Media;

namespace DynamicGeometry
{
    public partial class PointOnFigure : FreePoint, IPoint
    {
        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            Parameter = element.ReadDouble("Parameter");
            Recalculate();
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeDouble("Parameter", Parameter);
        }

        protected override Avalonia.Controls.Shapes.Shape CreateShape()
        {
            var result = Factory.CreateDependentPointShape();
            result.Fill = new SolidColorBrush(Color.FromArgb(255, 128, 255, 128));
            return result;
        }

        /// <summary>Where the point is along its figure. Setting it moves nothing by itself (a locus samples through it).</summary>
        public double Parameter { get; set; }

        /// <summary>The parameter as the grid edits it: the point goes there, and what is built on it follows</summary>
        [PropertyGridVisible]
        [PropertyGridName("Parameter")]
        public double ParameterDisplay
        {
            get
            {
                return Parameter;
            }
            set
            {
                Parameter = value;
                if (Drawing != null)
                {
                    this.RecalculateAllDependents();
                }
            }
        }

        public override double X => base.X;

        public override double Y => base.Y;

        public bool UseHitTestingForExistence = true;   // Broadens usefullness of PointOnFigure. Used in Tabula.

        public override bool AllowMove()
        {
            return !Locked;
        }

        public ILinearFigure LinearFigure
        {
            get
            {
                return (ILinearFigure)Dependencies.First();
            }
        }

        public override void MoveToCore(Point newPosition)
        {
            ILinearFigure figure = LinearFigure;
            Parameter = figure.GetNearestParameterFromPoint(newPosition);
            newPosition = figure.GetPointFromParameter(Parameter);
            base.MoveToCore(newPosition);
        }

        public override void Recalculate()
        {
            if (!Dependencies.Exists())
            {
                Exists = false;
                return;
            }

            var figure1 = LinearFigure;
            Point p = figure1.GetPointFromParameter(Parameter);
            bool HitTestFailed = (UseHitTestingForExistence) ? LinearFigure.HitTest(p) == null : false;
            if (!p.Exists() || HitTestFailed)
            {
                Exists = false;
                return;
            }

            Exists = true;
            Coordinates = p;
        }

        public static bool CanBeOnFigure(IFigure figure)
        {
            return figure is ILinearFigure;
        }

        /// <summary>Detaches the point from its figure (<see cref="PointSnapping.Release"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Free point")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void Release()
        {
            PointSnapping.Release(this);
        }
    }
}

