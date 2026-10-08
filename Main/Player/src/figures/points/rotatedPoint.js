// Port of Main/Avalonia/DynamicGeometry/Figures/Points/RotatedPoint.cs: a point turned about
// a center by an angle, which is a figure it depends on: a Number holding a typed value, or
// anything with an angle. Dependencies: the source, the center, the angle. Left out: the
// tying and untying of the angle (the editor's).

class RotatedPoint extends PointBase {
    get source() {
        return this.dependencies.length >= 1 && this.dependencies[0].isPoint === true ? this.dependencies[0] : null;
    }

    get rotationCenter() {
        return this.dependencies.length >= 2 && this.dependencies[1].isPoint === true ? this.dependencies[1] : null;
    }

    /** A Number or an angle provider */
    get angleSource() {
        return this.dependencies.length >= 3 ? this.dependencies[2] : null;
    }

    /** In degrees, counterclockwise */
    get angle() {
        const source = this.angleSource;
        if (source instanceof NumberFigure) {
            return source.value;
        }

        if (source != null && source.isAngleProvider === true) {
            return GeometryMath.toDegrees(source.angle);
        }

        return 0;
    }

    recalculate() {
        const source = this.source;
        const center = this.rotationCenter;
        if (source != null && center != null) {
            this.coordinates = GeometryMath.getRotationPoint(source.coordinates, center.coordinates, GeometryMath.toRadians(this.angle));
        }

        this.exists = allExist(this.dependencies) && this.coordinates.exists();
    }
}

FigureTypes.register("RotatedPoint", RotatedPoint);
