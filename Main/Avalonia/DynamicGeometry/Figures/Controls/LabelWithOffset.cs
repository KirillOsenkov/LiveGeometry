using Avalonia;
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
        /// <summary>
        /// A measurement depends on what it measures, and a dependent figure normally can't be
        /// dragged (its parents move instead). But all that dragging a measurement changes is the
        /// offset of the label from its anchor - so it can, like in the original DG.
        /// </summary>
        public override bool AllowMove()
        {
            return !Locked;
        }

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
