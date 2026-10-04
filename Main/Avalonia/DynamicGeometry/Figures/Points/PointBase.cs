using Avalonia;
using Avalonia.Controls.Shapes;
using System.Collections.Generic;

namespace DynamicGeometry
{
    public partial class PointBase : CoordinatesShapeBase<Shape>, IPoint
    {
        private static bool mSuppress;
        public static bool SuppressAutoLabelPoints
        {
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
            // a hidden helper (the square's midpoint) mustn't take a letter from the points on
            // screen: the square would be ABCE
            if (!Visible)
            {
                return base.GenerateFigureName(blacklist);
            }

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

        protected override string Kind
        {
            get
            {
                return "Point";
            }
        }

        /// <summary>A point goes by its name: "through E", not "through point E"</summary>
        public override string Noun
        {
            get
            {
                return null;
            }
        }

        // The label the point had last. A label is added and removed with the point and by
        // its Show name and Show coordinates, outside the undo history, so it must be the
        // same object every time: the history may hold a drag of it, and its place is its own.
        PointLabel keptLabel;

        // the point left the drawing with its label on (undo of adding it): it comes back with it
        bool returnsWithLabel;

        // new points are labeled once, when they are made: redo makes the point as it was
        bool wasInDrawing;

        public override void OnAddingToDrawing(Drawing drawing)
        {
            base.OnAddingToDrawing(drawing);

            // Make sure this is in the drawing's figure list before labeling.
            if (!Drawing.Figures.Contains(this))
            {
                return;
            }

            bool returning = wasInDrawing;
            bool withLabel = returnsWithLabel;
            wasInDrawing = true;
            returnsWithLabel = false;

            // while suppressed, whoever adds the point sees to its label (undo of a deletion
            // puts the label back itself)
            if (SuppressAutoLabelPoints)
            {
                return;
            }

            if (returning)
            {
                if (withLabel && Label == null && keptLabel != null && !Drawing.Figures.Contains(keptLabel))
                {
                    Label = keptLabel;
                    Drawing.Figures.Add(Label);
                }
            }
            else if (Settings.Instance.AutoLabelPoints && IsHitTestVisible && Visible)
            {
                ShowName = true;
            }
        }

        public override void OnRemovingFromDrawing(Drawing drawing)
        {
            if (Label != null)
            {
                keptLabel = Label;
                returnsWithLabel = true;
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
                keptLabel = Label;
                Drawing.Figures.Retire(Label);
                Label = null;
            }
        }

        private void AddLabelIfNecessary()
        {
            if (Label == null)
            {
                // the one it had, where it was, if it had one: saying nothing yet
                Label = keptLabel ?? Factory.CreatePointLabel(Drawing, new[] { this });
                Label.ShowName = false;
                Label.ShowCoordinates = false;

                // (one hidden by itself, in a file from before it couldn't be, shows again)
                Label.Visible = true;
                Drawing.Figures.Return(Label, owner: this);
            }
        }

        public PointLabel Label;
    }
}