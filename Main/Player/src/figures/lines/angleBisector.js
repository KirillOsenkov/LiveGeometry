// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/AngleBisector.cs: a ray halving an
// angle, built on its three points (vertex, side, side) or on an angle measurement. Left
// out: Convert to opposite angle.

class AngleBisector extends Ray {
    constructor() {
        super();
        this.coordinatesValue = new PointPair();

        /** Halves the angle under 180° between the sides, whichever way round they are */
        this.interior = true;

        /** Extends the bisector in both directions: the figure is then a line (isLine is the interface every line has) */
        this.wholeLine = false;

        /** Reads the sides the other way round, for the opposite angle */
        this.flipped = false;
    }

    get coordinates() {
        return this.coordinatesValue;
    }

    get angle() {
        const dependencies = this.getDependencies();
        let result = 0;
        if (dependencies != null) {
            result = this.flipped
                ? GeometryMath.toDegrees(GeometryMath.oAngle(dependencies[2].coordinates, dependencies[0].coordinates, dependencies[1].coordinates))
                : GeometryMath.toDegrees(GeometryMath.oAngle(dependencies[1].coordinates, dependencies[0].coordinates, dependencies[2].coordinates));
            if (this.interior && result > 180) {
                result = 360 - result;
            }
        }

        return result;
    }

    get onScreenCoordinates() {
        if (!this.wholeLine) {
            return super.onScreenCoordinates;
        }

        return GeometryMath.getLineFromSegment(this.coordinates, this.canvasLogicalBorders);
    }

    hitTest(point) {
        if (!this.wholeLine) {
            return super.hitTest(point);
        }

        const epsilon = this.toLogicalLength(this.strokeThickness) / 2 + this.cursorTolerance;
        return GeometryMath.isPointOnLine(this.coordinates, point, epsilon) ? this : null;
    }

    getNearestParameterFromPoint(point) {
        return this.wholeLine ? GeometryMath.getProjection(point, this.coordinates).ratio : super.getNearestParameterFromPoint(point);
    }

    getParameterDomain() {
        if (!this.wholeLine) {
            return super.getParameterDomain();
        }

        const coordinates = this.onScreenCoordinates;
        return [this.getNearestParameterFromPoint(coordinates.p1) * 2, this.getNearestParameterFromPoint(coordinates.p2) * 2];
    }

    readXml(element) {
        super.readXml(element);
        this.wholeLine = Xml.readBool(element, "Line", false);
        this.interior = Xml.readBool(element, "Interior", false);
    }

    /** The vertex and the two side points: the dependencies, or those of the angle measurement the bisector is built on */
    getDependencies() {
        let dependencies = this.dependencies;
        if (dependencies.length === 1) {
            const angle = dependencies[0];
            if (!(angle instanceof AngleMeasurement)) {
                return null;
            }

            dependencies = angle.dependencies;
        }

        return dependencies.length === 3 ? dependencies : null;
    }

    recalculate() {
        const dependencies = this.getDependencies();
        if (dependencies != null) {
            const vertex = dependencies[0].coordinates;
            const side1 = dependencies[this.flipped ? 2 : 1].coordinates;
            const side2 = dependencies[this.flipped ? 1 : 2].coordinates;

            // the halfway direction counterclockwise from side 1 to side 2; inside the angle
            // means the other way round when that sweep is the long way
            let halfway = GeometryMath.getAngleBisectorPoint(vertex, side1, side2);
            if (this.interior && halfway.exists() && GeometryMath.oAngle(side1, vertex, side2) > Math.PI) {
                halfway = vertex.minus(halfway.minus(vertex));
            }

            this.coordinatesValue = new PointPair(vertex, halfway);
            this.exists = allExist(this.dependencies) && allExist(dependencies) && halfway.exists();
        } else {
            this.exists = false;
        }
    }
}

FigureTypes.register("AngleBisector", AngleBisector);
