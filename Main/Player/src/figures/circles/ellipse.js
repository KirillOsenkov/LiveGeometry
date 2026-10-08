// Port of Main/Avalonia/DynamicGeometry/Figures/Circles/Ellipse.cs: center, end of the long
// axis, and a third point whose distance from the long axis is the short one

class Ellipse extends EllipseBase {
    get isShapeWithInterior() {
        return true;
    }

    get center() {
        return this.point(0);
    }

    get semiMajor() {
        return this.center.distance(this.point(1));
    }

    get semiMinor() {
        return GeometryMath.getDistanceToLine(this.point(2), new PointPair(this.center, this.point(1)));
    }

    get inclination() {
        return GeometryMath.getAngle(this.center, this.point(1));
    }
}

FigureTypes.register("Ellipse", Ellipse);
