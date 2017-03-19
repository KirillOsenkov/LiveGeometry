using System.Windows;
using System.Windows.Shapes;
using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry
{
    public partial class PointBase : CoordinatesShapeBase<Shape>, IPoint
    {
        private static bool mSuppress;
        public static bool SuppressAutoLabelPoints {
            get
            {
                return mSuppress;
            }
            set
            {
             mSuppress = value;   
            }
            }

        public override string GenerateFigureName(List<string> blacklist)
        {
            var alphabet = Settings.Instance.PointAlphabet;
            for (int i = 0; ; i++)
            {
                string number = i.ToString();
                if (i == 0)
                {
                    number = "";
                }
                foreach (var letter in alphabet)
                {
                    var candidate = letter.ToString() + number;
                    if (this.NameAvailable(candidate))
                    {
                        if (blacklist != null)
                        {
                            if (!blacklist.Contains(candidate))
                            {
                                return candidate;
                            }
                        }
                        else
                        {
                            return candidate;
                        }
                    }
                }
            }
        }

        public override void OnAddingToDrawing(Drawing drawing)
        {
            base.OnAddingToDrawing(drawing);

            // Make sure this is in the drawing's figure list before labeling.
            if (Settings.Instance.AutoLabelPoints && 
                         Drawing.Figures.Contains(this) && 
                         IsHitTestVisible &&
                         !SuppressAutoLabelPoints &&
                         Visible) ShowName = true;
        }

        public override void OnRemovingFromDrawing(Drawing drawing)
        {
            if (Label != null)
            {
                Drawing.Figures.Remove(Label);
                Label = null;
            }
        }

        protected override int DefaultZOrder()
        {
            return (int)ZOrder.Points;
        }

        protected override Shape CreateShape()
        {
            return Factory.CreatePointShape();
        }

        public virtual double X
        {
            get
            {
                return Coordinates.X;
            }
            set
            {
                Coordinates = Coordinates.SetX(value);
            }
        }

        public virtual double Y
        {
            get
            {
                return Coordinates.Y;
            }
            set
            {
                Coordinates = Coordinates.SetY(value);
            }
        }

        double PointSize
        {
            get
            {
                return Shape.ActualWidth / 2;
            }
        }

        public override IFigure HitTest(Point point)
        {
            double tolerance = CursorTolerance + ToLogical(PointSize);
            if (point.X >= Coordinates.X - tolerance
                && point.X <= Coordinates.X + tolerance
                && point.Y >= Coordinates.Y - tolerance
                && point.Y <= Coordinates.Y + tolerance)
            {
                return this;
            }

            return null;
        }

        [PropertyGridVisible]
        public override string Name
        {
            get
            {
                return base.Name;
            }
            set
            {
                base.Name = value;
                if (Label != null)
                {
                    Label.UpdateVisual();
                }
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Show name")]
        public bool ShowName
        {
            get
            {
                if (Label == null)
                {
                    return false;
                }
                return Label.ShowName;
            }
            set
            {
                if (ShowName == value)
                {
                    return;
                }
                AddLabelIfNecessary();
                Label.ShowName = value;
                RemoveLabelIfNecessary();
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Show coordinates")]
        public bool ShowCoordinates
        {
            get
            {
                if (Label == null)
                {
                    return false;
                }
                return Label.ShowCoordinates;
            }
            set
            {
                if (ShowCoordinates == value)
                {
                    return;
                }

                AddLabelIfNecessary();
                Label.ShowCoordinates = value;
                RemoveLabelIfNecessary();
            }
        }

        private void RemoveLabelIfNecessary()
        {
            if (Label != null && !Label.ShowName && !Label.ShowCoordinates)
            {
                Drawing.Figures.Remove(Label);
                Label = null;
            }
        }

        private void AddLabelIfNecessary()
        {
            if (Label == null)
            {
                Label = Factory.CreatePointLabel(Drawing, new [] { this });
                Label.ShowCoordinates = false;
                Drawing.Figures.Add(Label);
            }
        }

        public PointLabel Label;
    }
}