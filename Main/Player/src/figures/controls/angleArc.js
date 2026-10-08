// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/AngleArc.cs: the mark of an angle, an
// arc of a fixed size in pixels at the vertex (0 to 3 arcs, or the square sign at 90°).
// Left out: Convert to opposite angle.

class AngleArc extends CircleArc {
    static DefaultSize = 16;

    /** Gap between two arcs in pixels, on top of the stroke width */
    static ArcSpacing = 2;

    /** In radians: 0.005 degrees, which is when the label (two decimals) reads "90°" */
    static RightAngleTolerance = 0.005 * Math.PI / 180;

    constructor() {
        super();

        /** Radius of the (first) arc in pixels */
        this.size = AngleArc.DefaultSize;

        /** How the angle is marked: not at all, ), )) or ))) */
        this.arcCount = 1;
        this.shownSign = "Arc";
    }

    get isAngleArc() {
        return true;
    }

    readXml(element) {
        super.readXml(element);
        const arcs = element.getAttribute("Arcs");
        if (arcs != null && /^-?\d+$/.test(arcs.trim())) {
            this.arcCount = Math.max(0, Math.min(3, parseInt(arcs, 10)));
        }

        if (element.hasAttribute("Radius")) {
            this.size = Math.max(10, Math.min(100, Xml.readDouble(element, "Radius")));
        }
    }

    get radius() {
        return this.toLogicalLength(this.size);
    }

    get semiMajor() {
        return this.radius;
    }

    get semiMinor() {
        return this.radius;
    }

    get beginLocation() {
        return GeometryMath.scalePointBetweenTwo(this.center, this.point(1), this.radius / this.center.distance(this.point(1)));
    }

    /** The angle the mark is of, in degrees, like the number next to it */
    get measure() {
        return GeometryMath.toDegrees(this.angle);
    }

    /** An angle is two figures, the arc and the number, on the same vertex and the same two points: the other one of the pair, or null */
    static findCompanion(angleFigure) {
        const dependencies = angleFigure.dependencies;
        if (dependencies.length !== 3) {
            return null;
        }

        const wantArc = angleFigure instanceof AngleMeasurement;
        return dependencies[0].dependents.find(f =>
            f !== angleFigure
            && (wantArc ? f instanceof AngleArc : f instanceof AngleMeasurement)
            && f.dependencies.length === 3
            && f.dependencies[0] === dependencies[0]
            && f.dependencies.includes(dependencies[1])
            && f.dependencies.includes(dependencies[2])) ?? null;
    }

    /** Anywhere on the mark: from the first arc out to the last one, the gaps between them included */
    hitTest(point) {
        const found = super.hitTest(point);
        if (found != null || this.arcCount < 2 || this.shownSign !== "Arc") {
            return found;
        }

        const center = this.center;
        const distance = center.distance(point);
        const tolerance = this.cursorTolerance + this.logicalWidth() / 2;
        const outerRadius = this.toLogicalLength(this.size + (this.arcCount - 1) * (this.strokeThickness + AngleArc.ArcSpacing));
        if (distance < this.radius - tolerance || distance > outerRadius + tolerance) {
            return null;
        }

        const angleToPoint = GeometryMath.getAngle(center, point);
        return GeometryMath.isAngleBetweenAngles(angleToPoint, this.startAngle, this.endAngle, this.clockwise) ? this : null;
    }

    updateVisual() {
        const center = this.point(0);
        if (center.distance(this.point(1)) === 0 || center.distance(this.point(2)) === 0) {
            return;
        }

        // the way round it goes: the mirror image of a mark goes clockwise
        let angle = GeometryMath.oAngle(this.beginLocation, center, this.endLocation);
        if (this.clockwise && angle > 0) {
            angle = 2 * Math.PI - angle;
        }

        const isRightAngle = Math.abs(angle - Math.PI / 2) < AngleArc.RightAngleTolerance;
        this.shownSign = this.arcCount === 0 ? "None" : isRightAngle ? "RightAngle" : "Arc";
        this.shownAngle = angle;
    }

    updateExistence() {
        super.updateExistence();
        if (this.exists && !AngleArc.hasSides(this)) {
            this.exists = false;
        }
    }

    /** An angle (the arc or the number) exists only while both sides have a length */
    static hasSides(angleFigure) {
        const vertex = angleFigure.point(0);
        return vertex.distance(angleFigure.point(1)) !== 0 && vertex.distance(angleFigure.point(2)) !== 0;
    }

    render(renderer) {
        if (!this.isShown || this.shownSign === "None" || this.shownSign == null) {
            return;
        }

        const corner = this.toPhysical(this.point(0));
        const first = RightAngleMark.direction(corner, this.toPhysical(this.point(1)));
        const second = RightAngleMark.direction(corner, this.toPhysical(this.point(2)));
        if (first == null || second == null) {
            return;
        }

        const stroke = this.stroke;
        if (this.shownSign === "RightAngle") {
            // the school sign for 90 degrees: two sides of a little square, not an arc
            const points = RightAngleMark.getPoints(corner, first, second, RightAngleMark.Size);
            renderer.drawPolyline(points, stroke);
            return;
        }

        // the arcs: counterclockwise from the first side, each a stroke and a bit further out
        const startAngle = this.startAngle;
        const endAngle = this.endAngle;
        const fill = this.fill;
        for (let i = 0; i < this.arcCount; i++) {
            const radius = this.size + i * (this.strokeThickness + AngleArc.ArcSpacing);
            const start = corner.plus(first.scale(radius));
            const commands = [
                { op: "move", x: start.x, y: start.y },
                { op: "arc", cx: corner.x, cy: corner.y, rx: radius, ry: radius, rotation: 0, start: -startAngle, end: -endAngle, counterclockwise: !this.clockwise }
            ];
            renderer.drawPath(commands, stroke, i === 0 ? fill : null, new Rect(corner.x - radius, corner.y - radius, 2 * radius, 2 * radius));
        }
    }
}

FigureTypes.register("AngleArc", AngleArc);
