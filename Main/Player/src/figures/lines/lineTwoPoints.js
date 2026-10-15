// Port of Main/Avalonia/LiveGeometry/Figures/Lines/LineTwoPoints.cs

class LineTwoPoints extends LineBase {
    get isLine() {
        return true;
    }

    get onScreenCoordinates() {
        return GeometryMath.getLineFromSegment(this.coordinates, this.canvasLogicalBorders);
    }

    /** From the first point to the second, in degrees (what an expression reads) */
    get direction() {
        return DirectionConstraint.of(this.coordinates);
    }
}

FigureTypes.register("LineTwoPoints", LineTwoPoints);
