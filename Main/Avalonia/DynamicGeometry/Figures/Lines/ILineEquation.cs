using System.Xml;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public interface ILineEquation
    {
        PointPair LineCoordinates { get; }
        void Write(XmlWriter writer);
        void Recalculate();
    }

    public static class LineEquation
    {
        public static ILineEquation Read(LineByEquation parent, XElement element)
        {
            var m = element.ReadString("m");
            var b = element.ReadString("b");
            var A = element.ReadString("A");
            var B = element.ReadString("B");
            var C = element.ReadString("C");

            ILineEquation result = null;

            if (!m.IsEmpty() && !b.IsEmpty())
            {
                result = new SlopeInterseptLineEquation(parent, m, b);
            }
            else if (!A.IsEmpty() && !B.IsEmpty() && !C.IsEmpty())
            {
                result = new GeneralFormLineEquation(parent, A, B, C);
            }

            return result;
        }
    }

    public class SlopeInterseptLineEquation : ILineEquation
    {
        public SlopeInterseptLineEquation(IFigure parent, string slope, string intersept)
        {
            Slope = new DrawingExpression(parent) { Name = "m = " };
            Slope.Text = slope;
            Intersept = new DrawingExpression(parent) { Name = "b = " };
            Intersept.Text = intersept;
        }

        public PointPair LineCoordinates
        {
            get
            {
                if (Slope.Value == null || Intersept.Value == null)
                {
                    return new PointPair();
                }
                var m = Slope.Value();
                var b = Intersept.Value();

                if (m == 0)
                {
                    return new PointPair(0, b, 1, b);
                }
                else
                {
                    return new PointPair(0, b, 1, m + b);
                }
            }
        }

        public void Recalculate()
        {
            Slope.Recalculate();
            Intersept.Recalculate();
        }

        [PropertyGridVisible]
        [PropertyGridName("m = ")]
        public DrawingExpression Slope { get; private set; }

        [PropertyGridVisible]
        [PropertyGridName("b = ")]
        public DrawingExpression Intersept { get; private set; }

        public void Write(XmlWriter writer)
        {
            writer.WriteAttributeString("b", Intersept.Text);
            writer.WriteAttributeString("m", Slope.Text);
        }
    }

    public class GeneralFormLineEquation : ILineEquation
    {
        public GeneralFormLineEquation(IFigure parent, string a, string b, string c)
        {
            A = new DrawingExpression(parent) { Name = "A = " };
            A.Text = a;
            B = new DrawingExpression(parent) { Name = "B = " };
            B.Text = b;
            C = new DrawingExpression(parent) { Name = "C = " };
            C.Text = c;
        }

        public PointPair LineCoordinates
        {
            get
            {
                if (A.Value == null || B.Value == null || C.Value == null)
                {
                    return new PointPair();
                }
                var a = A.Value();
                var b = B.Value();
                var c = C.Value();

                if (a == 0 && b == 0)
                {
                    return new PointPair();
                }

                if (a == 0)
                {
                    var cb = -c / b;
                    return new PointPair(0, cb, 1, cb);
                }

                if (b == 0)
                {
                    var ca = -c / a;
                    return new PointPair(ca, 0, ca, 1);
                }

                return new PointPair(0, -c / b, -c / a, 0);
            }
        }

        public void Recalculate()
        {
            A.Recalculate();
            B.Recalculate();
            C.Recalculate();
        }

        [PropertyGridVisible]
        [PropertyGridName("A = ")]
        public DrawingExpression A { get; set; }

        [PropertyGridVisible]
        [PropertyGridName("B = ")]
        public DrawingExpression B { get; set; }

        [PropertyGridVisible]
        [PropertyGridName("C = ")]
        public DrawingExpression C { get; set; }

        public void Write(XmlWriter writer)
        {
            writer.WriteAttributeString("A", A.Text);
            writer.WriteAttributeString("B", B.Text);
            writer.WriteAttributeString("C", C.Text);
        }
    }
}
