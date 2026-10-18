// Port of Main/Avalonia/LiveGeometry/Figures/Controls/AnchoredLabel.cs: a label hung on one
// figure, its first dependency (a point's name, a line's or a circle's name, a text attached
// to a figure). It sits at its offset from an anchor on the figure, and a new one at the
// figure's default place (FigureAnchor); it keeps clear of other names (NameLayout), shows
// only while its figure does, and beside a line, which runs past the window, stays on screen.

class AnchoredLabel extends LabelWithOffset {
    constructor() {
        super();
        this.shown = true;
    }

    /** The figure the label hangs on */
    get figure() {
        return this.dependencies.length > 0 ? this.dependencies[0] : null;
    }

    /** Where the offset is measured from: on the figure, and for a line where the eye finds it */
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

    defaultLayer() {
        return ZOrder.PointLabels;
    }

    defaultOffset() {
        return FigureAnchor.defaultOffset(this.figure, this.measureSize(), this.drawing.coordinateSystem);
    }

    /** Whether the label goes while its figure is hidden or doesn't exist (a point's name doesn't) */
    get hidesWithFigure() {
        return true;
    }

    updateVisual() {
        const figure = this.figure;
        if (figure == null || this.drawing == null) {
            return;
        }

        super.updateVisual();
        if (!this.hidesWithFigure) {
            return;
        }

        let shown = this.visible && figure.visible && figure.exists;
        if (figure instanceof LineBase && !(figure instanceof Segment)) {
            // a line that runs outside the window has nowhere to write it
            shown = shown && FigureAnchor.isOnScreen(this.toPhysical(this.anchor), this.drawing.coordinateSystem);
            if (shown) {
                this.keepOnScreen();
            }
        }

        // a hidden figure keeps its label to itself
        this.shown = shown;
    }

    get isShown() {
        return super.isShown && this.shown;
    }

    /** The label pushed back in, whole, where a line's anchor by the edge of the window and the offset put it past the edge */
    keepOnScreen() {
        const topLeft = this.toPhysical(this.coordinates);
        const kept = FigureAnchor.keptOnScreen(topLeft, this.measureSize(), this.drawing.coordinateSystem);
        if (!kept.equals(topLeft) && kept.exists()) {
            this.coordinates = this.toLogical(kept);
        }
    }
}
