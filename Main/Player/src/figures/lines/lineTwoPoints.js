// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/LineTwoPoints.cs

class LineTwoPoints extends LineBase {
    get isLine() {
        return true;
    }

    get onScreenCoordinates() {
        return GeometryMath.getLineFromSegment(this.coordinates, this.canvasLogicalBorders);
    }
}

FigureTypes.register("LineTwoPoints", LineTwoPoints);
