// Port of Main/Avalonia/LiveGeometry/Figures/Controls/LabelWithOffset.cs: a label attached
// to something in the plane, at Offset pixels from its Anchor; and Measurement.

class LabelWithOffset extends LabelBase {
    constructor() {
        super();
        this.offsetValue = new Point();

        /** Whether the label has its place: the file's, one given to it, or the default one it got once */
        this.placed = false;

        /** Where the label is drawn instead of its place, while that is covered: from the anchor to its middle, in pixels (NameLayout) */
        this.nudge = null;

        // moved by a drag: drawn where it was put, whatever it covers, until it is next drawn
        this.movedByHand = false;

        // where the label was drawn last: when it changes, the labels drawn away from their place look again
        this.lastShownBox = null;
    }

    /** The label's top-left corner from its anchor, in pixels; set from anywhere, it is the label's place from then on */
    get offset() {
        return this.offsetValue;
    }

    set offset(value) {
        this.offsetValue = value;
        this.placed = true;
    }

    /** Where a new label goes, in pixels from the anchor; null for one that has no default place and keeps its offset */
    defaultOffset() {
        return null;
    }

    /** Whether the default place can be worked out now: a text to measure, a window with a size, no file being read */
    get canPlace() {
        return this.drawing != null
            && this.drawing.canvas != null
            && !this.drawing.isReading
            && this.drawing.coordinateSystem.physicalSize.x > 0
            && this.text !== "";
    }

    /** Gives the label its default place now, the nearest free one to it where the label keeps clear */
    place() {
        this.placed = true;
        const offset = this.defaultOffset();
        if (offset != null) {
            this.offset = this.keepsClear ? this.firstOffset(offset) : offset;
        }
    }

    /** Dragging the label changes only its offset from the anchor, not what it is built on */
    allowMove() {
        return !this.locked;
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

    /** Puts the label where its anchor and offset say, a new one at its default place; the text is the subclass's business */
    updateVisual() {
        if (!this.hasAnchor) {
            return;
        }

        if (!this.placed && this.canPlace) {
            this.place();
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
        // (the same place in other units: a label without one still gets the default)
        const unitLength = this.drawing.coordinateSystem.unitLength;
        this.offsetValue = new Point(this.offsetValue.x * unitLength, -this.offsetValue.y * unitLength);
    }

    readXml(element) {
        super.readXml(element);
        this.offset = new Point(Xml.readDouble(element, "OffsetX"), Xml.readDouble(element, "OffsetY"));

        // a file written by hand may leave the place out: then the label goes where a new one would
        this.placed = element.getAttribute("OffsetX") != null || element.getAttribute("OffsetY") != null;
    }
}

/** A label whose text is worked out from what it measures */
class Measurement extends LabelWithOffset {
}
