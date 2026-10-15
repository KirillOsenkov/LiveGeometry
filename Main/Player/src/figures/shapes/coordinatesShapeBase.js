// Port of Main/Avalonia/LiveGeometry/Figures/Shapes/CoordinatesShapeBase.cs: a shape at a place

class CoordinatesShapeBase extends ShapeBase {
    constructor() {
        super();
        this.coordinates = new Point();
    }

    /**
     * IMovable, as in C#: a shape at a place. Not every shape, which made the Drag tool take a
     * function graph with nothing to move for a figure with coordinates and throw, where the
     * press should have moved the view.
     */
    get isMovable() {
        return true;
    }

    moveToCore(newLocation) {
        this.coordinates = newLocation;
    }

    /** Where the figure is, for undo of a move: its coordinates, unless a subclass knows better */
    capturePlace() {
        return this.coordinates;
    }

    restorePlace(place) {
        this.moveTo(place);
    }

    get center() {
        return this.coordinates;
    }
}
