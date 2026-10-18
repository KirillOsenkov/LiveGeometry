// Port of Main/Avalonia/LiveGeometry/Figures/Controls/NameLabel.cs: the name written next to
// a figure, a point's (PointLabel) or a line's, a ray's, a segment's or a circle's
// (FigureLabel). Part of its figure (FigureBase.nameLabel), not a figure of its own.

class NameLabel extends AnchoredLabel {
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
}
