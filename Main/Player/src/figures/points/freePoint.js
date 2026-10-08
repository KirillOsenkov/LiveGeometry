// Port of Main/Avalonia/DynamicGeometry/Figures/Points/FreePoint.cs. Left out: the grid's
// verbs (snap to figure, convert to point by coordinates, the Bezier anchor verbs).

class FreePoint extends PointBase {
    readXml(element) {
        super.readXml(element);
        this.coordinates = new Point(Xml.readDouble(element, "X"), Xml.readDouble(element, "Y"));
    }

    /** Perf optimization: we know it exists, no need to call base */
    updateExistence() {
    }

    get x() {
        return super.x;
    }

    set x(value) {
        this.moveTo(new Point(value, this.y));
        this.recalculateAllDependents();
    }

    get y() {
        return super.y;
    }

    set y(value) {
        this.moveTo(new Point(this.x, value));
        this.recalculateAllDependents();
    }
}

FigureTypes.register("FreePoint", FreePoint);
