// Port of Main/Avalonia/LiveGeometry/Figures/Controls/AngleMeasurement.cs: the number of an
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
        // the text first: where a number that places itself goes depends on its size
        const text = NumberFormat.numberText(this.measure, this.decimalsToShow);
        this.setText(this.radians ? text + " rad" : text + "°");
        this.arrange();
        super.updateVisual();
    }

    /** Works out where the number goes before it is put there, when it places itself */
    arrange() {
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
        this.placedByHand = false;

        /** Where the layout put the number last, relative to its angle ({ turn, push }): it starts from there next time */
        this.layoutPlacement = null;
    }

    /** The measure of the angle the sweep chooses */
    get angle() {
        return AngleSweep.measure(this.sweep, super.angle);
    }

    readXml(element) {
        super.readXml(element);
        this.sweep = AngleSweep.read(element, AngleMeasurement.DefaultSweep);

        // (a file says where a number placed by hand goes; one without places itself, AngleLabelLayout)
        this.placedByHand = element.getAttribute("OffsetX") != null;
    }

    /** Whether the number goes where AngleLabelLayout puts it: one dragged stays where it was dragged to */
    get placesItself() {
        return !this.placedByHand;
    }

    moveToCore(newPosition) {
        this.placedByHand = true;
        super.moveToCore(newPosition);
    }

    capturePlace() {
        return { placedByHand: this.placedByHand, offset: this.offset };
    }

    restorePlace(place) {
        this.placedByHand = place.placedByHand;
        this.layoutPlacement = null;
        super.restorePlace(place.offset);
    }

    /** Whether the number is one the layout at the vertex keeps clear of or places: shown, of an angle there */
    takesPartInLayout(vertex) {
        return this.visible
            && this.exists
            && this.drawing != null
            && this.dependencies.length === 3
            && this.dependencies[0] === vertex
            && AngleArc.hasSides(this);
    }

    arrange() {
        if (this.visible && this.exists && this.canvas != null && this.drawing.coordinateSystem.physicalSize.x > 0 && AngleArc.hasSides(this)) {
            AngleLabelLayout.arrangeAt(this);
        }
    }

    /** Puts the number where the offset says, without working the layout out again (another number at the vertex just did) */
    moveToOffset() {
        this.coordinates = this.placeFromOffset();
    }

    /** The angle and the mark, as the layout takes them */
    layoutWedge() {
        const vertex = this.point(0);
        const counterclockwise = GeometryMath.oAngle(this.point(1), vertex, this.point(2));
        const measure = AngleSweep.measure(this.sweep, counterclockwise);
        return {
            vertex: this.toPhysical(vertex),
            start: GeometryMath.getAngle(vertex, this.point(1)),
            sweep: measure,
            isClockwise: AngleSweep.isClockwise(this.sweep, counterclockwise),
            clearance: this.markReach(measure),
            size: this.measureSize(),
            previous: this.layoutPlacement
        };
    }

    /** In pixels: how far from the vertex the mark reaches, its outer arc or the far corner of the square of a right angle */
    markReach(measure) {
        const arc = this.findArc();
        if (arc == null || !arc.visible || !arc.exists || arc.arcCount === 0) {
            return AngleLabelLayout.NoMarkClearance;
        }

        const stroke = arc.strokeThickness;
        if (Math.abs(measure - Math.PI / 2) < AngleArc.RightAngleTolerance) {
            return arc.size * Math.SQRT2 + stroke / 2;
        }

        return arc.size + (arc.arcCount - 1) * (stroke + AngleArc.ArcSpacing) + stroke / 2;
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
