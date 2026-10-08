// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/CoordinatesShapeBase.cs: a shape at a place

class CoordinatesShapeBase extends ShapeBase {
    constructor() {
        super();
        this.coordinates = new Point();
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
