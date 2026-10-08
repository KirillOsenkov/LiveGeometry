// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/LineAtAngle.cs: a line through a point at
// an angle to the x axis, the angle a figure it depends on (a Number, or anything with an
// angle). Dependencies: the point, then the angle. Left out: tying and untying.

class LineAtAngle extends LineBase {
    get isLine() {
        return true;
    }

    get angleSource() {
        return this.dependencies.length > 1 ? this.dependencies[1] : null;
    }

    get radians() {
        const source = this.angleSource;
        return source != null && source.isAngleProvider === true ? source.angle : NaN;
    }

    get coordinates() {
        const point = this.point(0);
        return new PointPair(point, GeometryMath.getTranslationPoint(point, 1, this.radians));
    }

    get onScreenCoordinates() {
        return GeometryMath.getLineFromSegment(this.coordinates, this.canvasLogicalBorders);
    }

    /** An angle that isn't a number leaves no line */
    updateExistence() {
        super.updateExistence();
        if (this.exists && !this.coordinates.p2.exists()) {
            this.exists = false;
        }
    }

    /** In degrees */
    get angle() {
        const source = this.angleSource;
        return source instanceof NumberFigure ? source.value : GeometryMath.toDegrees(this.radians);
    }
}

FigureTypes.register("LineAtAngle", LineAtAngle);
