// Port of Main/Avalonia/LiveGeometry/Figures/Controls/LabelWithOffset.cs: a label attached
// to something in the plane, at Offset pixels from its Anchor; and Measurement.

class LabelWithOffset extends LabelBase {
    constructor() {
        super();

        /** The label's top-left corner from its anchor, in pixels */
        this.offset = new Point();

        /** Where the label is drawn instead of its place, while that is covered: from the anchor to its middle, in pixels (NameLayout) */
        this.nudge = null;

        // moved by a drag: drawn where it was put, whatever it covers, until it is next drawn
        this.movedByHand = false;

        // where the label was drawn last: when it changes, the labels drawn away from their place look again
        this.lastShownBox = null;
    }

    /** Whether the label is drawn elsewhere when something covers its place (NameLayout) */
    get keepsClear() {
        return false;
    }

    /** The point the label belongs to, which its other places keep clear of */
    get ownPoint() {
        return null;
    }

    /** The lines and circles that may run through the anchor */
    figuresThroughAnchor() {
        return [];
    }

    /** Where the label is drawn on the canvas, in pixels; null while it isn't */
    shownBox() {
        if (this.drawing == null || this.drawing.canvas == null || !this.isShown) {
            return null;
        }

        const size = this.measureSize();
        const topLeft = this.toPhysical(this.coordinates);
        return size.width > 0 && size.height > 0 && topLeft.exists() ? Rect.fromPoints(topLeft, size) : null;
    }

    /** The offset a new label takes, given its default one: the nearest free place to it (NameLayout.firstPlace) */
    firstOffset(offset) {
        if (this.drawing == null || this.drawing.canvas == null || this.drawing.isReading) {
            return offset;
        }

        const anchor = this.toPhysical(this.anchor);
        if (!anchor.exists()) {
            return offset;
        }

        return NameLayout.firstPlace(this, anchor.plus(offset)).minus(anchor);
    }

    onRemovingFromDrawing(drawing) {
        NameLayout.forget(this, drawing);
        super.onRemovingFromDrawing(drawing);
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
        this.movedByHand = true;
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

    /** Whether there is an anchor to measure from: something the label is built on, unless a subclass takes its anchor from elsewhere */
    get hasAnchor() {
        return this.dependencies.length > 0;
    }

    /** Puts the label where its anchor and offset say; the text is the subclass's business */
    updateVisual() {
        if (!this.hasAnchor) {
            return;
        }

        this.coordinates = this.placeFromOffset();
        if (this.keepsClear) {
            this.keepClear();
        }

        const shown = this.shownBox();
        if (!(shown == null ? this.lastShownBox == null : shown.equals(this.lastShownBox))) {
            this.lastShownBox = shown;
            NameLayout.recheck(this);
        }
    }

    /** Drawn elsewhere while something covers the label's place, unless it was just put there by hand */
    keepClear() {
        if (this.movedByHand || this.drawing == null || this.drawing.canvas == null || this.drawing.isReading) {
            this.movedByHand = false;
            NameLayout.forget(this, this.drawing);
            return;
        }

        this.coordinates = this.toLogical(NameLayout.place(this, this.toPhysical(this.coordinates)));
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
