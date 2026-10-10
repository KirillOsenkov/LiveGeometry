// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/AngleArc.cs: the mark of an angle, an
// arc of a fixed size in pixels at the vertex (0 to 3 arcs, or the square sign at 90°), of
// the angle its sweep chooses (AngleSweep; the number next to it says the same one).

class AngleArc extends CircleArc {
    static DefaultSize = 20;

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

    /** A new mark says the angle under 180°, whichever way round its sides were clicked */
    get defaultSweep() {
        return AngleSweep.Smaller;
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

    /** Whether the place is in what a filled mark fills: the sector under the arcs, or the square of a right angle */
    isInsideFill(point) {
        const center = this.center;
        if (this.shownSign === "RightAngle") {
            const corner = this.toPhysical(center);
            const first = RightAngleMark.direction(corner, this.toPhysical(this.point(1)));
            const second = RightAngleMark.direction(corner, this.toPhysical(this.point(2)));
            if (first == null || second == null) {
                return false;
            }

            // (the two directions are at a right angle: the square's own coordinates)
            const offset = this.toPhysical(point).minus(corner);
            const along = offset.x * first.x + offset.y * first.y;
            const across = offset.x * second.x + offset.y * second.y;
            return along >= 0 && along <= this.size && across >= 0 && across <= this.size;
        }

        if (this.shownSign !== "Arc" || center.distance(point) > this.radius) {
            return false;
        }

        return GeometryMath.isAngleBetweenAngles(GeometryMath.getAngle(center, point), this.startAngle, this.endAngle, this.isClockwise);
    }

    /** Anywhere on the mark: from the first arc out to the last one, the gaps between them included; inside a filled mark */
    hitTest(point) {
        if (this.fill != null && this.visible && this.isInsideFill(point)) {
            return this;
        }

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
        return GeometryMath.isAngleBetweenAngles(angleToPoint, this.startAngle, this.endAngle, this.isClockwise) ? this : null;
    }

    updateVisual() {
        const center = this.point(0);
        if (center.distance(this.point(1)) === 0 || center.distance(this.point(2)) === 0) {
            return;
        }

        // the angle the sweep chooses
        const angle = this.angle;
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

    /** An angle (the arc or the number) exists only while its sides have a length: every point after the vertex, two or one (an angle to the x axis) */
    static hasSides(angleFigure) {
        const count = angleFigure.dependencies.length;
        if (count < 2) {
            return false;
        }

        const vertex = angleFigure.point(0);
        for (let i = 1; i < count; i++) {
            if (vertex.distance(angleFigure.point(i)) === 0) {
                return false;
            }
        }

        return true;
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

        // a style with a fill fills the angle under the mark - the square of the sign, the
        // sector under the first arc - filled and not outlined: the sides are drawn by the
        // lines through the vertex, and the arcs are never filled themselves
        const stroke = this.stroke;
        const fill = this.fill;
        if (this.shownSign === "RightAngle") {
            // the school sign for 90 degrees: two sides of a little square, not an arc, as
            // big as the arc would be (the radius is the square's side)
            const points = RightAngleMark.getPoints(corner, first, second, this.size);
            if (fill != null) {
                renderer.drawPolygon([corner, ...points], null, fill, true);
            }

            renderer.drawPolyline(points, stroke);
            return;
        }

        // the arcs: from the first side the way the sweep goes, each a stroke and a bit further out
        const startAngle = this.startAngle;
        const endAngle = this.endAngle;
        const counterclockwise = !this.isClockwise;
        if (fill != null) {
            const radius = this.size;
            const start = corner.plus(first.scale(radius));
            const sector = [
                { op: "move", x: corner.x, y: corner.y },
                { op: "line", x: start.x, y: start.y },
                { op: "arc", cx: corner.x, cy: corner.y, rx: radius, ry: radius, rotation: 0, start: -startAngle, end: -endAngle, counterclockwise },
                { op: "close" }
            ];
            renderer.drawPath(sector, null, fill, new Rect(corner.x - radius, corner.y - radius, 2 * radius, 2 * radius));
        }

        for (let i = 0; i < this.arcCount; i++) {
            const radius = this.size + i * (this.strokeThickness + AngleArc.ArcSpacing);
            const start = corner.plus(first.scale(radius));
            const commands = [
                { op: "move", x: start.x, y: start.y },
                { op: "arc", cx: corner.x, cy: corner.y, rx: radius, ry: radius, rotation: 0, start: -startAngle, end: -endAngle, counterclockwise }
            ];
            renderer.drawPath(commands, stroke, null, new Rect(corner.x - radius, corner.y - radius, 2 * radius, 2 * radius));
        }
    }
}

FigureTypes.register("AngleArc", AngleArc);
