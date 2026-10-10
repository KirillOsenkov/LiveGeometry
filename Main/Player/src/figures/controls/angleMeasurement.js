// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/AngleMeasurement.cs: the number of an
// angle at its vertex, in degrees or radians, and the angle of a segment to the x axis.

class AngleMeasurementBase extends Measurement {
    constructor() {
        super();
        this.radiansValue = false;
    }

    get isAngleProvider() {
        return true;
    }

    get radians() {
        return this.radiansValue;
    }

    set radians(value) {
        this.radiansValue = value;
        this.updateVisual();
    }

    /** The number shown: the angle in the unit chosen */
    get measure() {
        const angle = this.angle;
        return this.radians ? angle : GeometryMath.toDegrees(angle);
    }

    /** The angle in radians, 0 to 2π: counterclockwise from the first side to the second */
    get angle() {
        return GeometryMath.oAngle(this.point(1), this.point(0), this.point(2));
    }

    get anchor() {
        return this.point(0);
    }

    updateVisual() {
        super.updateVisual();
        const text = NumberFormat.toString(GeometryMath.round(this.measure, this.decimalsToShow));
        this.setText(this.radians ? text + " rad" : text + "°");
    }

    readXml(element) {
        super.readXml(element);
        this.radiansValue = Xml.readBool(element, "Radians", false);
    }
}

class AngleMeasurement extends AngleMeasurementBase {
    /** A new angle is the one under 180°, whichever way round its sides were clicked */
    static DefaultSweep = AngleSweep.Smaller;

    constructor() {
        super();

        /** Which of the two angles at the vertex the number says (AngleSweep); the mark next to it shows the same one */
        this.sweep = AngleMeasurement.DefaultSweep;
    }

    /** The measure of the angle the sweep chooses */
    get angle() {
        return AngleSweep.measure(this.sweep, super.angle);
    }

    readXml(element) {
        super.readXml(element);
        this.sweep = AngleSweep.read(element, AngleMeasurement.DefaultSweep);
    }

    /** No angle while a side has no length */
    updateExistence() {
        super.updateExistence();
        if (this.exists && !AngleArc.hasSides(this)) {
            this.exists = false;
        }
    }

    /** The arc that was created together with this label: same vertex, same two sides; null if deleted */
    findArc() {
        return AngleArc.findCompanion(this);
    }

    get arcCount() {
        const arc = this.findArc();
        return arc != null ? arc.arcCount : 0;
    }
}

class HorizontalAngleMeasurement extends AngleMeasurementBase {
    /** The angle of the segment to the x axis, counterclockwise, in radians (the base reads a third point) */
    get angle() {
        return GeometryMath.oHAngle(this.point(0), this.point(1));
    }

    get measure() {
        return this.radians ? this.angle : GeometryMath.toDegrees(this.angle);
    }

    updateVisual() {
        if (this.dependencies.length === 0) {
            return;
        }

        this.coordinates = this.placeFromOffset();
        this.setText(NumberFormat.toDegreeString(GeometryMath.toDegrees(this.angle)));
    }
}

FigureTypes.register("AngleMeasurement", AngleMeasurement);
FigureTypes.register("HorizontalAngleMeasurement", HorizontalAngleMeasurement);
