// Port of Main/Avalonia/DynamicGeometry/Figures/Circles/Circle.cs

class Circle extends CircleBase {
    get isShapeWithInterior() {
        return true;
    }

    get center() {
        return this.point(0);
    }

    get radius() {
        return this.center.distance(this.point(1));
    }

    get inclination() {
        return GeometryMath.getAngle(this.center, this.point(1));
    }
}

FigureTypes.register("Circle", Circle);
