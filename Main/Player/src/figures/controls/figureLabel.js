// Port of Main/Avalonia/LiveGeometry/Figures/Controls/FigureLabel.cs: the name of a line,
// ray, segment or circle, written next to it, at Offset pixels from an anchor on it.

class FigureLabel extends Measurement {
    constructor() {
        super();
        this.placed = false;
        this.shown = true;
    }

    get figure() {
        return this.dependencies.length > 0 ? this.dependencies[0] : null;
    }

    onAddingToDrawing(drawing) {
        super.onAddingToDrawing(drawing);
        const figure = this.figure;
        if (figure instanceof FigureBase && figure.nameLabel == null) {
            figure.nameLabel = this;
        }
    }

    onRemovingFromDrawing(drawing) {
        super.onRemovingFromDrawing(drawing);
        const figure = this.figure;
        if (figure instanceof FigureBase && figure.nameLabel === this) {
            figure.nameLabel = null;
        }
    }

    defaultLayer() {
        return ZOrder.PointLabels;
    }

    /** Where the offset is measured from: on the figure, and for a line where the eye finds it (FigureAnchor) */
    get anchor() {
        return FigureAnchor.of(this.figure, this.drawing?.coordinateSystem);
    }

    get keepsClear() {
        return true;
    }

    /** A segment's name doesn't go across its segment, a circle's across its circle */
    figuresThroughAnchor() {
        return this.figure != null ? [this.figure] : [];
    }

    updateVisual() {
        const figure = this.figure;
        if (figure == null || this.drawing == null) {
            return;
        }

        this.setText(NameDisplay.format(figure.name));
        if (!this.placed && this.text !== "") {
            this.placed = true;
            this.offset = this.firstOffset(FigureAnchor.defaultOffset(figure, this.measureSize(), this.drawing.coordinateSystem));
        }

        super.updateVisual();
        let shown = this.visible && figure.visible;
        if (figure instanceof LineBase && !(figure instanceof Segment)) {
            shown = shown && FigureAnchor.isOnScreen(this.toPhysical(this.anchor), this.drawing.coordinateSystem);
            if (shown) {
                this.keepOnScreen();
            }
        }

        this.shown = shown;
    }

    get isShown() {
        return super.isShown && this.shown;
    }

    /** The name pushed back in, whole, where a line's anchor by the edge of the window and the offset put it past the edge */
    keepOnScreen() {
        const topLeft = this.toPhysical(this.coordinates);
        const kept = FigureAnchor.keptOnScreen(topLeft, this.measureSize(), this.drawing.coordinateSystem);
        if (!kept.equals(topLeft) && kept.exists()) {
            this.coordinates = this.toLogical(kept);
        }
    }

    readXml(element) {
        super.readXml(element);
        this.placed = true;
    }
}

FigureTypes.register("FigureLabel", FigureLabel);
