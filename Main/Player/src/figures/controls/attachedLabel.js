// Port of Main/Avalonia/LiveGeometry/Figures/Controls/AttachedLabel.cs: a text attached to a
// figure, written next to a point, a segment, a line, a ray or a circle as its name would be
// (FigureAnchor) and moving with it, but saying whatever is typed. Plain text: no [...]
// parts. It keeps clear as a name does (NameLayout).

class AttachedLabel extends LabelWithOffset {
    constructor() {
        super();
        this.placed = false;
        this.shown = true;
    }

    /** The figure the text is attached to */
    get figure() {
        return this.dependencies.length > 0 ? this.dependencies[0] : null;
    }

    get anchor() {
        return FigureAnchor.of(this.figure, this.drawing?.coordinateSystem);
    }

    get keepsClear() {
        return true;
    }

    get ownPoint() {
        const figure = this.figure;
        return figure != null && figure.isPoint === true ? figure : null;
    }

    /** A point's lines and circles; a segment, a line or a circle itself */
    figuresThroughAnchor() {
        const figure = this.figure;
        if (figure == null) {
            return [];
        }

        return figure.isPoint === true ? [...figure.dependents, ...figure.dependencies] : [figure];
    }

    /** Dragged by itself: that changes its offset from the anchor, not the figure */
    allowMove() {
        return !this.locked;
    }

    defaultLayer() {
        return ZOrder.PointLabels;
    }

    updateVisual() {
        const figure = this.figure;
        if (figure == null || this.drawing == null) {
            return;
        }

        if (!this.placed && this.text !== "") {
            // placed for good once the drawing is there to keep clear of
            this.offset = FigureAnchor.defaultOffset(figure, this.measureSize(), this.drawing.coordinateSystem);
            if (this.drawing.canvas != null && !this.drawing.isReading) {
                this.placed = true;
                this.offset = this.firstOffset(this.offset);
            }
        }

        super.updateVisual();
        let shown = this.visible && figure.visible && figure.exists;
        if (figure instanceof LineBase && !(figure instanceof Segment)) {
            shown = shown && FigureAnchor.isOnScreen(this.toPhysical(this.anchor), this.drawing.coordinateSystem);
            if (shown) {
                const topLeft = this.toPhysical(this.coordinates);
                const kept = FigureAnchor.keptOnScreen(topLeft, this.measureSize(), this.drawing.coordinateSystem);
                if (!kept.equals(topLeft) && kept.exists()) {
                    this.coordinates = this.toLogical(kept);
                }
            }
        }

        this.shown = shown;
    }

    get isShown() {
        return super.isShown && this.shown;
    }

    readXml(element) {
        super.readXml(element);
        this.setText(Label.unescape(element.getAttribute("Text") ?? ""));

        // a file written by hand may leave the place out: then the text goes where a new one would
        this.placed = element.getAttribute("OffsetX") != null || element.getAttribute("OffsetY") != null;
    }
}

FigureTypes.register("AttachedLabel", AttachedLabel);
