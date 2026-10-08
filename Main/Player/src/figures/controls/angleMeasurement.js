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

    get measure() {
        const measure = GeometryMath.oAngle(this.point(1), this.point(0), this.point(2));
        return this.radians ? measure : GeometryMath.toDegrees(measure);
    }

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
    moveToCore(newPosition) {
        super.moveToCore(newPosition.plus(0.2));
    }

    updateVisual() {
        if (this.dependencies.length === 0) {
            return;
        }

        this.coordinates = this.placeFromOffset();
        this.setText(NumberFormat.toDegreeString(GeometryMath.toDegrees(GeometryMath.oHAngle(this.point(0), this.point(1)))));
    }
}

FigureTypes.register("AngleMeasurement", AngleMeasurement);
FigureTypes.register("HorizontalAngleMeasurement", HorizontalAngleMeasurement);
