// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/LabelWithOffset.cs: a label attached
// to something in the plane, at Offset pixels from its Anchor; and Measurement.

class LabelWithOffset extends LabelBase {
    constructor() {
        super();

        /** The label's top-left corner from its anchor, in pixels */
        this.offset = new Point();
    }

    /** What the offset is measured from, in the plane */
    get anchor() {
        return new Point();
    }

    /** Where the anchor and the offset put the label right now, in the plane */
    placeFromOffset() {
        return this.toLogical(this.toPhysical(this.anchor).plus(this.offset));
    }

    /** Moving the label changes its offset from the anchor */
    moveToCore(newPosition) {
        this.offset = this.toPhysical(newPosition).minus(this.toPhysical(this.anchor));
        super.moveToCore(newPosition);
    }

    capturePlace() {
        return this.offset;
    }

    restorePlace(place) {
        this.offset = place;
        if (this.drawing != null) {
            this.updateVisual();
        }
    }

    /** Puts the label where its anchor and offset say; the text is the subclass's business */
    updateVisual() {
        if (this.dependencies.length === 0) {
            return;
        }

        this.coordinates = this.placeFromOffset();
    }

    /** A drawing from before version 1 stored the offset in units of the plane */
    upgradeOffsetFromUnits() {
        const unitLength = this.drawing.coordinateSystem.unitLength;
        this.offset = new Point(this.offset.x * unitLength, -this.offset.y * unitLength);
    }

    readXml(element) {
        super.readXml(element);
        this.offset = new Point(Xml.readDouble(element, "OffsetX"), Xml.readDouble(element, "OffsetY"));
    }
}

/** A label whose text is worked out from what it measures; draggable, since that only changes its offset */
class Measurement extends LabelWithOffset {
    allowMove() {
        return !this.locked;
    }
}
