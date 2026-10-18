// Port of Main/Avalonia/LiveGeometry/Figures/Points/PointBase.cs. Left out: auto-labeling
// of new points, the kept label across undo (a point's label comes from the file).

class PointBase extends CoordinatesShapeBase {
    get isPoint() {
        return true;
    }

    /** A hidden helper mustn't take a letter from the points on screen (GenerateFigureName) */
    generateFigureName() {
        if (!this.visible) {
            return super.generateFigureName();
        }

        const alphabet = Settings.pointAlphabet;
        for (let i = 0; ; i++) {
            const number = i === 0 ? "" : String(i);
            for (const letter of alphabet) {
                const candidate = letter + number;
                if (this.nameAvailable(candidate)) {
                    return candidate;
                }
            }
        }
    }

    onRemovingFromDrawing(drawing) {
        const label = this.nameLabel;
        if (label != null) {
            this.nameLabel = null;
            drawing.figures.remove(label);
        }
    }

    /** A constructed point: dragged, it moves what it is built on (FreePoint and the others that can be dragged by themselves say otherwise) */
    defaultLayer() {
        return ZOrder.FixedPoints;
    }

    get x() {
        return this.coordinates.x;
    }

    set x(value) {
        this.coordinates = this.coordinates.withX(value);
    }

    get y() {
        return this.coordinates.y;
    }

    set y(value) {
        this.coordinates = this.coordinates.withY(value);
    }

    /** The marker's size in pixels (Shape.ActualWidth), from the style on screen */
    get pointSize() {
        return this.resolvedStyle?.size ?? 10;
    }

    hitTest(point) {
        const tolerance = this.cursorTolerance + this.toLogicalLength(this.pointSize / 2);
        if (point.x >= this.coordinates.x - tolerance
            && point.x <= this.coordinates.x + tolerance
            && point.y >= this.coordinates.y - tolerance
            && point.y <= this.coordinates.y + tolerance) {
            return this;
        }

        return null;
    }

    get showName() {
        return this.nameLabel instanceof PointLabel && this.nameLabel.showName;
    }

    get showCoordinates() {
        return this.nameLabel instanceof PointLabel && this.nameLabel.showCoordinates;
    }

    render(renderer) {
        if (!this.isShown || !this.coordinates.exists()) {
            return;
        }

        const style = this.resolvedStyle;
        if (style == null) {
            return;
        }

        renderer.drawMarker(
            this.toPhysical(this.coordinates),
            style.size,
            style.shape ?? PointShape.Circle,
            style.character ?? null,
            this.stroke,
            style.fillBrush !== undefined ? style.fillBrush : null);
    }
}
