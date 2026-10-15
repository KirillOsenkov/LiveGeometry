// Port of Main/Avalonia/LiveGeometry/Figures/Shapes/Polygon.cs

class Polygon extends PolygonBase {
    get isPolygon() {
        return true;
    }

    get numberOfSides() {
        return this.dependencies.length;
    }
}

FigureTypes.register("Polygon", Polygon);
