using System.Windows;
using System.Xml;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public abstract class LabelWithOffset : LabelBase, IMovable
    {
        public Point Offset { get; set; }

        public override void ReadXml(XElement element)
        {
            base.ReadXml(element);
            Offset = new Point(element.ReadDouble("OffsetX"), element.ReadDouble("OffsetY"));
        }

        public override void WriteXml(XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeDouble("OffsetX", Offset.X);
            writer.WriteAttributeDouble("OffsetY", Offset.Y);
        }


    }

    public abstract class Measurement : LabelWithOffset
    {
        //private int mDecimalsToShow = 2;
        //[PropertyGridName("Decimals (0-10)")]
        //[PropertyGridVisible]
        //public virtual int DecimalsToShow
        //{
        //    get
        //    {
        //        return mDecimalsToShow;
        //    }
        //    set
        //    {
        //        if (value >= 0 && value <= 10)
        //        {
        //            mDecimalsToShow = value;
        //            UpdateVisual();
        //        }
        //    }
        //}

        //public override void ReadXml(XElement element)
        //{
        //    base.ReadXml(element);
        //    DecimalsToShow = (int)element.ReadDouble("DecimalsToShow");
        //}

        //public override void WriteXml(XmlWriter writer)
        //{
        //    base.WriteXml(writer);
        //    writer.WriteAttributeDouble("DecimalsToShow", (double)DecimalsToShow);
        //}
    }
}
