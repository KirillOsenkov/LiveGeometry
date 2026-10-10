// Port of Main/Avalonia/DynamicGeometry/Figures/Circles/ArcBase.cs: arcs, segments and sectors
// of ellipses and circles. The begin and end are where the rays from the center to the
// begin and end points meet the ellipse.

class EllipseArcBase extends ShapeBase {
    constructor() {
        super();
        this.sweepValue = this.defaultSweep;
    }

    get isLinearFigure() {
        return true;
    }

    get isEllipse() {
        return true;
    }

    get isArc() {
        return true;
    }

    get isAngleProvider() {
        return true;
    }

    get isLengthProvider() {
        return true;
    }

    /** Whether the arc is closed by its chord (a segment) and whether by two radii (a sector) */
    get isSegmentShape() {
        return false;
    }

    get isSectorShape() {
        return false;
    }

    get semiMajor() {
        return this.point(0).distance(this.point(1));
    }

    get semiMinor() {
        // as for Ellipse: the third point's distance from the long axis
        return GeometryMath.getDistanceToLine(this.point(2), new PointPair(this.point(0), this.point(1)));
    }

    logicalWidth() {
        return this.toLogicalLength(this.strokeThickness);
    }

    /** Which of the two arcs between the begin and the end this is (AngleSweep) */
    get sweep() {
        return this.sweepValue;
    }

    set sweep(value) {
        if (this.sweepValue !== value) {
            this.sweepValue = value;
            if (this.drawing != null) {
                this.recalculateAllDependents();
            }
        }
    }

    /** What a new one is: the way round the tool's clicks went (an angle's mark says the smaller angle) */
    get defaultSweep() {
        return AngleSweep.Counterclockwise;
    }

    /** The counterclockwise angle from the begin to the end, 0 to 2π: what the sweep chooses from */
    get counterclockwiseAngle() {
        return GeometryMath.oAngle(this.beginLocation, this.center, this.endLocation);
    }

    /** Whether the arc goes clockwise from its begin to its end right now */
    get isClockwise() {
        return AngleSweep.isClockwise(this.sweep, this.counterclockwiseAngle);
    }

    /** The length of the arc: Simpson's rule over the parameter */
    get length() {
        const start = this.parametricAngle(this.isClockwise ? this.endLocation : this.beginLocation);
        return GeometryMath.ellipseArcLength(this.semiMajor, this.semiMinor, start, this.parametricSweep);
    }

    /** The point half way along the arc (where a perimeter's number sits) */
    get arcMiddle() {
        const begin = this.parametricAngle(this.beginLocation);
        const sweep = this.parametricSweep;
        return this.pointAtParametricAngle(begin + (this.isClockwise ? -sweep : sweep) / 2);
    }

    updateVisual() {
    }

    getNearestParameterFromPoint(point) {
        let result = GeometryMath.getAngle(this.center, point);
        const clockwise = this.isClockwise;
        let a1 = clockwise ? this.endAngle : this.startAngle;
        let a2 = clockwise ? this.startAngle : this.endAngle;
        if (!Settings.pointsOnEllipticalsUseAbsoluteAngle) {
            const inclination = this.inclination;
            result -= inclination;
            a1 -= inclination;
            a2 -= inclination;
        }

        // off the arc: its nearer end, the short way round
        const onArc = a2 < a1 ? result <= a2 || result >= a1 : result >= a1 && result <= a2;
        if (!onArc) {
            const toStart = Math.abs(GeometryMath.ieeeRemainder(result - a1, 2 * Math.PI));
            const toEnd = Math.abs(GeometryMath.ieeeRemainder(result - a2, 2 * Math.PI));
            result = toStart <= toEnd ? a1 : a2;
        }

        if (this.flipped) {
            result = -result;
        }

        return result;
    }

    getPointFromParameter(parameter) {
        const center = this.center;
        let angleToPoint = parameter;
        if (this.flipped) {
            angleToPoint = -angleToPoint;
        }

        if (!Settings.pointsOnEllipticalsUseAbsoluteAngle) {
            angleToPoint += this.inclination;
        }

        const intersections = GeometryMath.getIntersectionOfEllipseFigureAndLine(this, new PointPair(center, GeometryMath.getTranslationPoint(center, 1, angleToPoint)));
        const direction1 = GeometryMath.getAngle(center, intersections.p1);
        const cDiff = Math.cos(angleToPoint) - Math.cos(direction1);
        const sDiff = Math.sin(angleToPoint) - Math.sin(direction1);
        return isWithinEpsilon(cDiff) && isWithinEpsilon(sDiff) ? intersections.p1 : intersections.p2;
    }

    hitTest(point) {
        // the fill (a pixel hit test in C#: here the region the arc encloses)
        const style = this.resolvedStyle;
        if (style != null && style.isFilled === true && this.isInsideFill(point)) {
            return this;
        }

        // the edge
        const width = this.logicalWidth();
        const angleToPoint = GeometryMath.getAngle(this.center, point);
        if (GeometryMath.isAngleBetweenAngles(angleToPoint, this.startAngle, this.endAngle, this.isClockwise)) {
            const fromEdge = GeometryMath.radialDistanceToEllipse(this.center, this.semiMajor, this.semiMinor, this.inclination, point);
            if (Math.abs(fromEdge) < this.cursorTolerance + width / 2) {
                return this;
            }
        }

        const epsilon = this.toLogicalLength(this.strokeThickness) / 2 + this.cursorTolerance;

        // the chord
        if (this.isSegmentShape && GeometryMath.isPointOnSegment(new PointPair(this.beginLocation, this.endLocation), point, epsilon)) {
            return this;
        }

        // the radii
        if (this.isSectorShape
            && (GeometryMath.isPointOnSegment(new PointPair(this.center, this.beginLocation), point, epsilon)
                || GeometryMath.isPointOnSegment(new PointPair(this.center, this.endLocation), point, epsilon))) {
            return this;
        }

        return null;
    }

    /** The region the figure fills, as Avalonia fills its path: the arc closed by its chord, or by the radii for a sector */
    isInsideFill(point) {
        const outline = this.sampleArc(48);
        if (outline.length < 2) {
            return false;
        }

        if (this.isSectorShape) {
            outline.push(this.center);
        }

        return GeometryMath.isPointInPolygon(outline, point);
    }

    /** Points along the arc from its begin to its end, logical */
    sampleArc(count) {
        const a = this.semiMajor;
        const b = this.semiMinor;
        if (!(a > 0 && b > 0)) {
            return [];
        }

        const begin = this.parametricAngle(this.beginLocation);
        let sweep = this.parametricSweep;
        if (this.isClockwise) {
            sweep = -sweep;
        }

        const result = [];
        for (let i = 0; i <= count; i++) {
            result.push(this.pointAtParametricAngle(begin + sweep * i / count));
        }

        return result;
    }

    pointAtParametricAngle(t) {
        return GeometryMath.pointOnEllipse(this.center, this.semiMajor, this.semiMinor, this.inclination, t);
    }

    getParameterDomain() {
        const clockwise = this.isClockwise;
        const a1 = clockwise ? this.endAngle : this.startAngle;
        let a2 = clockwise ? this.startAngle : this.endAngle;
        if (a2 < a1) {
            a2 += 2 * Math.PI;
        }

        return [a1, a2];
    }

    get inclination() {
        return GeometryMath.getAngle(this.center, this.point(1));
    }

    get beginPointIndex() {
        return 3;
    }

    get endPointIndex() {
        return 4;
    }

    /** Where the ray from the center through the point meets the ellipse */
    locationTowards(point) {
        const center = this.center;
        const intersections = GeometryMath.getIntersectionOfEllipseAndLine(center, this.semiMajor, this.semiMinor, this.inclination, new PointPair(center, point));
        const i1 = intersections.p1.distance(point);
        const i2 = intersections.p2.distance(point);
        return i1 < i2 ? intersections.p1 : intersections.p2;
    }

    get beginLocation() {
        return this.locationTowards(this.point(this.beginPointIndex));
    }

    get endLocation() {
        return this.locationTowards(this.point(this.endPointIndex));
    }

    get startAngle() {
        return GeometryMath.getAngle(this.center, this.beginLocation);
    }

    get endAngle() {
        return GeometryMath.getAngle(this.center, this.endLocation);
    }

    /** The central angle of the arc, 0 to 2π: the measure of the region the sweep chooses */
    get angle() {
        return AngleSweep.measure(this.sweep, this.counterclockwiseAngle);
    }

    /** How far the arc goes around, 0 to 2π, in the angle t that parametrizes its ellipse */
    get parametricSweep() {
        const begin = this.parametricAngle(this.beginLocation);
        const end = this.parametricAngle(this.endLocation);
        let sweep = this.isClockwise ? begin - end : end - begin;
        if (sweep < 0) {
            sweep += 2 * Math.PI;
        }

        return sweep;
    }

    parametricAngle(point) {
        const center = this.center;
        const inclination = this.inclination;
        const cos = Math.cos(inclination);
        const sin = Math.sin(inclination);
        const dx = point.x - center.x;
        const dy = point.y - center.y;
        const along = dx * cos + dy * sin;
        const across = -dx * sin + dy * cos;
        return Math.atan2(across / this.semiMinor, along / this.semiMajor);
    }

    /** The area between the arc and the two radii to its ends: a * b * t / 2 */
    get sectorArea() {
        const a = this.semiMajor;
        const b = this.semiMinor;
        return a > 0 && b > 0 ? a * b * this.parametricSweep / 2 : 0;
    }

    /** The area between the arc and its chord: a * b * (t - sin t) / 2 */
    get segmentArea() {
        const a = this.semiMajor;
        const b = this.semiMinor;
        if (!(a > 0 && b > 0)) {
            return 0;
        }

        const sweep = this.parametricSweep;
        return a * b * (sweep - Math.sin(sweep)) / 2;
    }

    get center() {
        return this.point(0);
    }

    readXml(element) {
        super.readXml(element);
        this.sweepValue = AngleSweep.read(element, this.defaultSweep);
    }

    /** The path in pixels: the arc, closed by the chord or the radii where the figure is */
    buildPath() {
        const center = this.toPhysical(this.center);
        const rx = this.toPhysicalLength(this.semiMajor);
        const ry = this.toPhysicalLength(this.semiMinor);
        if (!center.exists() || !(rx > 0) || !(ry > 0)) {
            return null;
        }

        // canvas angles run the other way (y down): a counterclockwise sweep here is one
        // of decreasing canvas angle
        const clockwise = this.isClockwise;
        const begin = this.parametricAngle(this.beginLocation);
        const sweep = this.parametricSweep;
        const end = clockwise ? begin - sweep : begin + sweep;
        const commands = [];
        const start = this.toPhysical(this.beginLocation);
        if (this.isSectorShape) {
            commands.push({ op: "move", x: center.x, y: center.y });
            commands.push({ op: "line", x: start.x, y: start.y });
        } else {
            commands.push({ op: "move", x: start.x, y: start.y });
        }

        commands.push({ op: "arc", cx: center.x, cy: center.y, rx, ry, rotation: -this.inclination, start: -begin, end: -end, counterclockwise: !clockwise });
        if (this.isSectorShape || this.isSegmentShape) {
            commands.push({ op: "close" });
        }

        return commands;
    }

    render(renderer) {
        if (!this.isShown) {
            return;
        }

        const commands = this.buildPath();
        if (commands == null) {
            return;
        }

        const reach = this.toPhysicalLength(Math.max(this.semiMajor, this.semiMinor));
        const center = this.toPhysical(this.center);
        renderer.drawPath(commands, this.stroke, this.fill, new Rect(center.x - reach, center.y - reach, 2 * reach, 2 * reach));
    }
}

class CircleArcBase extends EllipseArcBase {
    get isCircle() {
        return true;
    }

    get beginLocation() {
        return this.point(this.beginPointIndex);
    }

    get beginPointIndex() {
        return 1;
    }

    get endLocation() {
        return GeometryMath.scalePointBetweenTwo(this.center, this.point(2), this.radius / this.center.distance(this.point(2)));
    }

    get endPointIndex() {
        return 2;
    }

    /** No arc while its radius is 0 or its end is on the center */
    updateExistence() {
        super.updateExistence();
        if (this.exists && this.dependencies.length > 2 && (!(this.radius > 0) || !(this.center.distance(this.point(2)) > 0))) {
            this.exists = false;
        }
    }

    get length() {
        return this.radius * this.angle;
    }

    get radius() {
        return this.semiMajor;
    }

    get semiMinor() {
        return this.semiMajor;
    }

    get inclination() {
        return GeometryMath.getAngle(this.center, this.point(1));
    }
}
