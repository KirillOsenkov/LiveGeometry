using Avalonia;

namespace DynamicGeometry
{
    public class CircleByEquation : CircleBase, ICircle, IShapeWithInterior, IRenamableExpressions, IExpressionOwner
    {
        [PropertyGridVisible(false)]
        public System.Collections.Generic.IEnumerable<DrawingExpression> Expressions
        {
            get
            {
                return new[] { X, Y, R };
            }
        }

        public void RenameInExpressions(ExpressionRenamer renamer)
        {
            X?.RenameInExpression(renamer);
            Y?.RenameInExpression(renamer);
            R?.RenameInExpression(renamer);
        }

        public void RebindExpressions()
        {
            X?.Rebind();
            Y?.Rebind();
            R?.Rebind();
        }

        [PropertyGridVisible(false)]
        public System.Collections.Generic.IReadOnlyList<string> ExpressionTexts
        {
            get
            {
                return new[] { X?.Text, Y?.Text, R?.Text };
            }
            set
            {
                if (X != null && Y != null && R != null)
                {
                    X.Text = value[0];
                    Y.Text = value[1];
                    R.Text = value[2];
                }
            }
        }

        [PropertyGridVisible(false)]
        public override Point Center
        {
            get 
            {
                if (X.Value == null || Y.Value == null)
                {
                    return new Point();
                }

                return new Point(X.Value(), Y.Value());
            }
        }

        [PropertyGridVisible(false)]
        public override double Radius
        {
            get 
            {
                if (R.Value == null)
                {
                    return 0;
                }
                var radius = R.Value();
                return radius > 0 ? radius : Math.Epsilon;
            }
        }

        /// <summary>
        /// A center or a radius without a value (an expression that doesn't compile, a radius
        /// below 0 or undefined) leaves no circle, as for a circle by radius: it was a dot at
        /// (0, 0), or at the center, with what was built on it
        /// </summary>
        public override void UpdateExistence()
        {
            base.UpdateExistence();
            if (Exists
                && (X?.Value == null
                    || Y?.Value == null
                    || R?.Value == null
                    || !Center.Exists()
                    || !(R.Value() >= 0)))
            {
                Exists = false;
            }
        }

        /// <summary>"with center (1, A.Y) and radius 2"</summary>
        public override string Construction
        {
            get
            {
                if (X == null || Y == null || R == null)
                {
                    return null;
                }

                return "with center (" + X.Text + ", " + Y.Text + ") and radius " + R.Text;
            }
        }

        [PropertyGridVisible]
        public DrawingExpression X { get; set; }

        [PropertyGridVisible]
        public DrawingExpression Y { get; set; }

        [PropertyGridVisible]
        public DrawingExpression R { get; set; }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            X = new DrawingExpression(this, "Center X =", element.ReadString("X"));
            Y = new DrawingExpression(this, "Center Y =", element.ReadString("Y"));
            R = new DrawingExpression(this, "Radius =", element.ReadString("R"));
            base.ReadXml(element);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeString("X", X.Text);
            writer.WriteAttributeString("Y", Y.Text);
            writer.WriteAttributeString("R", R.Text);
        }
    }
}