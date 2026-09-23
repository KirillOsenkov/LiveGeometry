using Avalonia;
using System.Xml;
using System.Xml.Linq;

namespace DynamicGeometry
{
    /// <summary>
    /// A label attached to something in the plane: it sits at <see cref="Offset"/> from its
    /// <see cref="Anchor"/>.
    /// </summary>
    public abstract class LabelWithOffset : LabelBase, IMovable
    {
        /// <summary>
        /// The label's top-left corner from its anchor, in pixels: what the label is attached
        /// to zooms, the text doesn't, and the gap between them shouldn't either (in units of
        /// the plane it shrank onto the point when zoomed out and drifted away when zoomed in).
        /// Files before version 1 stored it in units; see <see cref="UpgradeOffsetFromUnits"/>.
        /// </summary>
        public Point Offset { get; set; }

        /// <summary>
        /// What the offset is measured from, in the plane: the point, the middle of the
        /// segment, the vertex of the angle
        /// </summary>
        public abstract Point Anchor { get; }

        /// <summary>Where the anchor and the offset put the label right now, in the plane</summary>
        protected Point PlaceFromOffset()
        {
            return ToLogical(ToPhysical(Anchor).Plus(Offset));
        }

        /// <summary>Moving the label changes its offset from the anchor</summary>
        public override void MoveToCore(Point newPosition)
        {
            Offset = ToPhysical(newPosition).Minus(ToPhysical(Anchor));
            base.MoveToCore(newPosition);
        }

        /// <summary>Puts the label where its anchor and offset say; the text is the subclass's business</summary>
        public override void UpdateVisual()
        {
            if (Dependencies.IsEmpty())
            {
                return;
            }

            Coordinates = PlaceFromOffset();
            base.UpdateVisual();
        }

        /// <summary>
        /// A drawing from before version 1 stored the offset in units of the plane; this turns
        /// it into pixels at the zoom the drawing opened at, the best guess for the zoom it was
        /// placed at.
        /// </summary>
        public void UpgradeOffsetFromUnits()
        {
            double unitLength = Drawing.CoordinateSystem.UnitLength;
            Offset = new Point(Offset.X * unitLength, -Offset.Y * unitLength);
        }

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
    }
}
