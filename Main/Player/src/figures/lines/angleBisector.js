// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/AngleBisector.cs: a ray halving an
// angle, built on its three points (vertex, side, side) or on an angle measurement, the
// angle its sweep chooses (AngleSweep).

class AngleBisector extends Ray {
    /** A new bisector halves the angle under 180°, whichever way round its sides were clicked */
    static DefaultSweep = AngleSweep.Smaller;

    constructor() {
        super();
        this.coordinatesValue = new PointPair();

        /** Which of the two angles between the sides is halved; a bisector of a measurement halves the angle that says */
        this.sweepValue = AngleBisector.DefaultSweep;

        /** Extends the bisector in both directions: the figure is then a line (isLine is the interface every line has) */
        this.wholeLine = false;
    }

    get coordinates() {
        return this.coordinatesValue;
    }

    get sweep() {
        const dependencies = this.dependencies;
        return dependencies.length === 1 && dependencies[0] instanceof AngleMeasurement ? dependencies[0].sweep : this.sweepValue;
    }

    /** The angle halved, in degrees */
    get angle() {
        const dependencies = this.getDependencies();
        return dependencies != null ? GeometryMath.toDegrees(AngleSweep.measure(this.sweep, AngleBisector.counterclockwiseAngle(dependencies))) : 0;
    }

    /** The counterclockwise angle from the first side to the second, 0 to 2π: what the sweep chooses from */
    static counterclockwiseAngle(dependencies) {
        return GeometryMath.oAngle(dependencies[1].coordinates, dependencies[0].coordinates, dependencies[2].coordinates);
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
        this.sweepValue = AngleSweep.read(element, AngleBisector.DefaultSweep);
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
            const side1 = dependencies[1].coordinates;
            const side2 = dependencies[2].coordinates;

            // the halfway direction counterclockwise from side 1 to side 2, or the other
            // way round when the sweep chooses the clockwise region
            let halfway = GeometryMath.getAngleBisectorPoint(vertex, side1, side2);
            if (halfway.exists() && AngleSweep.isClockwise(this.sweep, AngleBisector.counterclockwiseAngle(dependencies))) {
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
