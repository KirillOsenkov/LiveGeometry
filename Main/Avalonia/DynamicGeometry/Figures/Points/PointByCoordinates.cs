namespace DynamicGeometry
{
    public class PointByCoordinates : PointBase, IPoint, IRenamableExpressions, IConditionalProperties, IExpressionOwner
    {
        public System.Collections.Generic.IEnumerable<DrawingExpression> Expressions
        {
            get
            {
                return new[] { XExpression, YExpression };
            }
        }

        public PointByCoordinates()
        {
            XExpression = new DrawingExpression(this) { Name = "X = " };
            YExpression = new DrawingExpression(this) { Name = "Y = " };
        }

        public void RenameInExpressions(ExpressionRenamer renamer)
        {
            XExpression.RenameInExpression(renamer);
            YExpression.RenameInExpression(renamer);
        }

        public void RebindExpressions()
        {
            XExpression.Rebind();
            YExpression.Rebind();
        }

        public System.Collections.Generic.IReadOnlyList<string> ExpressionTexts
        {
            get
            {
                return new[] { XExpression.Text, YExpression.Text };
            }
            set
            {
                XExpression.Text = value[0];
                YExpression.Text = value[1];
            }
        }

        /// <summary>"at (2, A.Y + 1)"</summary>
        public override string Construction
        {
            get
            {
                return "at (" + XExpression.Text + ", " + YExpression.Text + ")";
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("X = ")]
        public DrawingExpression XExpression { get; private set; }
        
        [PropertyGridVisible]
        [PropertyGridName("Y = ")]
        public DrawingExpression YExpression { get; private set; }

        /// <summary>Lets go of the coordinates: a free point where it is (<see cref="PointSnapping.Release"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Free point")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void Release()
        {
            if (PointSnapping.CanFree(this))
            {
                PointSnapping.Release(this);
            }
        }

        /// <summary>
        /// The point is where its X and Y say, also when they are plain numbers and it
        /// depends on nothing: a drag moved it nowhere, and left an undo step that undid
        /// nothing. Like a locked point, it also keeps what is built on it from being dragged.
        /// </summary>
        public override bool AllowMove()
        {
            return false;
        }

        public bool CanEdit(string propertyName)
        {
            return propertyName != nameof(Release) || PointSnapping.CanFree(this);
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

        public override void Recalculate()
        {
            // an expression that doesn't compile (it names a figure that isn't there) gives no
            // place: the point is nowhere, not at (0, 0) with everything built on it
            if (XExpression == null
                || XExpression.Value == null
                || YExpression == null
                || YExpression.Value == null)
            {
                Exists = false;
                return;
            }
            Coordinates = new Avalonia.Point(XExpression.Value(), YExpression.Value());

            // (see ReflectedPoint: X = A.X has no value while A is not there)
            Exists = Dependencies.Exists() && Coordinates.Exists();
        }

        public override void OnAddingToDrawing(Drawing drawing)
        {
            // Recalculate in order to compile expressions and have accurate coordinates.  
            // Needed when autoLabelPoints is on. -D.H.
            Recalculate();
            base.OnAddingToDrawing(drawing);
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            XExpression.Text = element.ReadString("X");
            YExpression.Text = element.ReadString("Y");
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeString("X", XExpression.Text);
            writer.WriteAttributeString("Y", YExpression.Text);
        }
    }
}
