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

            // Where the file says, until the drawing is worked out (it is, as each figure
            // comes in). Not worked out here: what the point is on may be built on figures
            // that are nowhere yet, and a point that doesn't exist kept the (0, 0) it got
            // from that - and wrote it into the next file.
            Parameter = element.ReadDouble("Parameter");
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
            if (!p.Exists())
            {
                Exists = false;
                return;
            }

            // Where the parameter says, also when that place is off the figure (past the
            // end of an arc that turned away) and the point doesn't exist: its coordinates
            // go into the file, and left where the point last existed - or where a locus
            // last sampled it - the same drawing was saved with other numbers each time.
            Coordinates = p;

            // A graph is hit by its samples, and those end at the edges of the window: a
            // point on it is wherever the function has a value. (Hit-tested, it stopped
            // existing, with everything built on it, when the view was panned away from it.)
            // A locus likewise: its point is worked out exactly, and may be a hair off the
            // curve as drawn through its samples.
            bool hitTestFailed = UseHitTestingForExistence
                && !(figure1 is FunctionGraph)
                && !(figure1 is Locus)
                && LinearFigure.HitTest(p) == null;
            Exists = !hitTestFailed;
        }

        /// <summary>
        /// Whether a click puts a new point on the figure. Not on the mark of an angle,
        /// though it is an arc: the mark is a sign of a fixed size in pixels, not a figure
        /// of the plane - a click near the vertex glued the point to it, and the point
        /// moved with every zoom.
        /// </summary>
        public static bool CanBeOnFigure(IFigure figure)
        {
            return figure is ILinearFigure && !(figure is AngleArc);
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

