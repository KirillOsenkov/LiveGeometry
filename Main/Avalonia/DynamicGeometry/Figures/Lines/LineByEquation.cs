namespace DynamicGeometry
{
    public class LineByEquation : LineBase, ILine, IRenamableExpressions, IExpressionOwner
    {
        [PropertyGridVisible(false)]
        public System.Collections.Generic.IEnumerable<DrawingExpression> Expressions
        {
            get
            {
                return Equation != null ? Equation.Expressions : new DrawingExpression[0];
            }
        }

        public void RenameInExpressions(ExpressionRenamer renamer)
        {
            if (Equation != null)
            {
                foreach (var expression in Equation.Expressions)
                {
                    expression.RenameInExpression(renamer);
                }
            }
        }

        public void RebindExpressions()
        {
            if (Equation != null)
            {
                foreach (var expression in Equation.Expressions)
                {
                    expression.Rebind();
                }
            }
        }

        [PropertyGridVisible(false)]
        public System.Collections.Generic.IReadOnlyList<string> ExpressionTexts
        {
            get
            {
                return System.Linq.Enumerable.ToArray(System.Linq.Enumerable.Select(Expressions, expression => expression.Text));
            }
            set
            {
                int index = 0;
                foreach (var expression in Expressions)
                {
                    expression.Text = value[index++];
                }
            }
        }

        protected override string Kind
        {
            get
            {
                return "Line";
            }
        }

        /// <summary>"y = 2x + 1", "2x + 3y - 6 = 0": the equation as its expressions say it</summary>
        public override string Construction
        {
            get
            {
                switch (Equation)
                {
                    case SlopeInterseptLineEquation slope:
                        return "y = " + ConstructionText.Sum((slope.Slope.Text, "x"), (slope.Intersept.Text, ""));
                    case GeneralFormLineEquation general:
                        return ConstructionText.Sum((general.A.Text, "x"), (general.B.Text, "y"), (general.C.Text, "")) + " = 0";
                    default:
                        return null;
                }
            }
        }

        public override PointPair OnScreenCoordinates
        {
            get
            {
                return Math.GetLineFromSegment(Coordinates, CanvasLogicalBorders);
            }
        }

        public override PointPair Coordinates
        {
            get
            {
                return Equation.LineCoordinates;
            }
        }

        /// <summary>
        /// An equation without a value - an expression that doesn't compile, A = B = 0, a slope
        /// of sqrt(-1) - leaves no line: it was a line through (0, 0) going nowhere, with
        /// what was built on it
        /// </summary>
        public override void UpdateExistence()
        {
            base.UpdateExistence();
            if (!Exists)
            {
                return;
            }

            var coordinates = Equation?.LineCoordinates;
            if (coordinates == null
                || !coordinates.Value.P1.Exists()
                || !coordinates.Value.P2.Exists()
                || coordinates.Value.P1 == coordinates.Value.P2)
            {
                Exists = false;
            }
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            Equation = LineEquation.Read(this, element);
            base.ReadXml(element);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            Equation.Write(writer);
        }

        [PropertyGridComplexTypeState(ComplexTypeState.Expanded)]
        [PropertyGridVisible]
        [PropertyGridName("Equation")]
        public ILineEquation Equation { get; set; }
    }
}
