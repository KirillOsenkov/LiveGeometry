// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/RightAngleMark.cs: the little square
// in the corner that says "these two are perpendicular". Not a figure: a perpendicular line
// owns one and shows it by itself. Which corner it sits in is remembered and never derived
// from the geometry again. Left out: the click that moves it to the next corner.

class RightAngleMark {
    /** Side of the square in pixels; an angle measurement uses the same */
    static Size = 10;

    static Stroke = Color.fromArgb(0xA0, 0x70, 0x70, 0x70);

    static CornerAttribute = "RightAngleCorner";
    static VisibleAttribute = "RightAngleMark";

    constructor(owner) {
        this.owner = owner;
        this.cornerValue = 0;
        this.isEnabled = true;

        /** The three points of the sign in pixels while it shows; null while hidden */
        this.points = null;
    }

    /** 0 to 3, counterclockwise. 0 is between the direction of the base line and that direction turned left by 90 degrees. */
    get corner() {
        return this.cornerValue;
    }

    set corner(value) {
        this.cornerValue = ((value % 4) + 4) % 4;
    }

    hide() {
        this.points = null;
    }

    /**
     * vertex: where the two lines meet, logical. baseLine: the line the owner is perpendicular
     * to, logical; its direction is what the corner counts from. baseFigure: the figure drawn
     * along the base line, if any: a mark whose side along it would reach past its end isn't drawn.
     */
    show(drawing, vertex, baseLine, baseFigure) {
        const coordinateSystem = drawing.coordinateSystem;
        let along = baseLine.p2.minus(baseLine.p1);
        let across = new Point(-along.y, along.x);
        switch (this.corner) {
            case 1:
                along = along.negate();
                break;
            case 2:
                along = along.negate();
                across = across.negate();
                break;
            case 3:
                across = across.negate();
                break;
        }

        const cornerPoint = coordinateSystem.toPhysical(vertex);
        const first = RightAngleMark.direction(cornerPoint, coordinateSystem.toPhysical(vertex.plus(along)));
        const second = RightAngleMark.direction(cornerPoint, coordinateSystem.toPhysical(vertex.plus(across)));
        if (!this.isEnabled || first == null || second == null || RightAngleMark.hasAngleMeasuredAt(drawing, vertex)) {
            this.hide();
            return;
        }

        const points = RightAngleMark.getPoints(cornerPoint, first, second, RightAngleMark.Size);
        if (baseFigure != null && baseFigure.hitTest(coordinateSystem.toLogical(points[0])) == null) {
            this.hide();
            return;
        }

        this.points = points;
    }

    render(renderer) {
        if (this.points == null) {
            return;
        }

        renderer.drawPolyline(this.points, { color: RightAngleMark.Stroke, width: 1, dash: null, join: "miter" });
    }

    /** The corner with the most room, to start in: towards the farther end of the base line and towards the given point across it */
    static getRoomiestCorner(vertex, baseLine, pointAcross) {
        const along = baseLine.p2.minus(baseLine.p1);
        const across = new Point(-along.y, along.x);
        const forward = vertex.distance(baseLine.p2) >= vertex.distance(baseLine.p1);
        const toPoint = pointAcross.minus(vertex);
        const left = toPoint.x * across.x + toPoint.y * across.y >= 0;
        if (left) {
            return forward ? 0 : 1;
        }

        return forward ? 3 : 2;
    }

    /** The three points of the sign: out along one side, the far corner of the square, back onto the other side */
    static getPoints(corner, firstDirection, secondDirection, size) {
        return [
            corner.plus(firstDirection.scale(size)),
            corner.plus(firstDirection.plus(secondDirection).scale(size)),
            corner.plus(secondDirection.scale(size))
        ];
    }

    /** The unit vector from one physical point to another; null if they coincide */
    static direction(from, to) {
        const length = from.distance(to);
        if (!(length > 1e-9) || !Number.isFinite(length)) {
            return null;
        }

        return to.minus(from).scale(1 / length);
    }

    static hasAngleMeasuredAt(drawing, vertex) {
        const tolerance = drawing.coordinateSystem.cursorTolerance;
        return drawing.figures.list.some(a => a.isAngleArc === true && a.visible && a.exists && a.center.distance(vertex) < tolerance);
    }

    /** Null if the file says nothing about the corner (it is older than the mark) */
    readXml(element) {
        this.isEnabled = Xml.readBool(element, RightAngleMark.VisibleAttribute, true);
        const cornerAttribute = element.getAttribute(RightAngleMark.CornerAttribute);
        if (cornerAttribute != null && /^-?\d+$/.test(cornerAttribute.trim())) {
            this.corner = parseInt(cornerAttribute, 10);
            return this.corner;
        }

        return null;
    }
}
