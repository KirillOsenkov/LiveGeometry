// Port of Main/Avalonia/LiveGeometry/Figures/Regions/BoundaryPiece.cs and RegionBoundary.cs:
// the boundary of a set made of other regions by an operation (difference, intersection,
// union), traced from the operands' boundaries. Every operand's boundary is cut into pieces
// at the crossings with the others, and a piece is kept when the result lies on one side of
// it and not the other, judged at its middle. The kept pieces, the result on their left,
// give the area (Green's theorem, exact for segments and arcs), the perimeter, the loops to
// draw, the centroid and the nearest point of the edge.

const RegionOperation = {
    Difference: "Difference",
    Intersection: "Intersection",
    Union: "Union"
};

/** A stretch of a boundary: straight or an arc of an ellipse, parameter 0 at its start to 1 at its end */
class BoundaryPiece {
    get middle() {
        return this.pointAt(0.5);
    }

    /** The unit normal to the left of the direction of travel at the parameter */
    leftNormalAt(t) {
        const direction = this.directionAt(t);
        return new Point(-direction.y, direction.x);
    }

    /** Whether the two pieces are one and the same stretch, run either way */
    isSameAs(other, tolerance) {
        if (!RegionBoundary.near(this.middle, other.middle, tolerance)) {
            return false;
        }

        return RegionBoundary.near(this.start, other.start, tolerance) && RegionBoundary.near(this.end, other.end, tolerance)
            || RegionBoundary.near(this.start, other.end, tolerance) && RegionBoundary.near(this.end, other.start, tolerance);
    }
}

class SegmentPiece extends BoundaryPiece {
    constructor(start, end) {
        super();
        this.start = start;
        this.end = end;
    }

    pointAt(t) {
        return new Point(this.start.x + (this.end.x - this.start.x) * t, this.start.y + (this.end.y - this.start.y) * t);
    }

    directionAt(t) {
        const length = this.length;
        return length > 0 ? new Point((this.end.x - this.start.x) / length, (this.end.y - this.start.y) / length) : new Point(1, 0);
    }

    get length() {
        return this.start.distance(this.end);
    }

    /** The piece's share of the area to its left: the integral of (x - ox) dy - (y - oy) dx, halved */
    areaTerm(origin) {
        return ((this.start.x - origin.x) * (this.end.y - origin.y) - (this.end.x - origin.x) * (this.start.y - origin.y)) / 2;
    }

    part(t0, t1) {
        return new SegmentPiece(this.pointAt(t0), this.pointAt(t1));
    }

    reversed() {
        return new SegmentPiece(this.end, this.start);
    }

    nearestPoint(point) {
        return GeometryMath.getNearestPointOnSegment(point, new PointPair(this.start, this.end));
    }

    get bounds() {
        const left = Math.min(this.start.x, this.end.x);
        const top = Math.min(this.start.y, this.end.y);
        return new Rect(left, top, Math.abs(this.end.x - this.start.x), Math.abs(this.end.y - this.start.y));
    }

    addSamples(points) {
        points.push(this.end);
    }

    /** Where the point lies along the segment as a parameter (0 at the start, 1 at the end, beyond either outside) */
    parameterOf(point) {
        const dx = this.end.x - this.start.x;
        const dy = this.end.y - this.start.y;
        const lengthSquared = dx * dx + dy * dy;
        return lengthSquared > 0 ? ((point.x - this.start.x) * dx + (point.y - this.start.y) * dy) / lengthSquared : 0;
    }
}

/** An arc of an ellipse from the parametric angle from to to; clockwise when to is below from */
class ArcPiece extends BoundaryPiece {
    constructor(center, semiMajor, semiMinor, inclination, from, to) {
        super();
        this.center = center;
        this.semiMajor = semiMajor;
        this.semiMinor = semiMinor;
        this.inclination = inclination;
        this.from = from;
        this.to = to;
    }

    get sweep() {
        return this.to - this.from;
    }

    get isCounterclockwise() {
        return this.to >= this.from;
    }

    get isCircle() {
        return this.semiMajor === this.semiMinor && this.inclination === 0;
    }

    get start() {
        return this.at(this.from);
    }

    get end() {
        return this.at(this.to);
    }

    at(theta) {
        return GeometryMath.pointOnEllipse(this.center, this.semiMajor, this.semiMinor, this.inclination, theta);
    }

    pointAt(t) {
        return this.at(this.from + (this.to - this.from) * t);
    }

    directionAt(t) {
        const theta = this.from + (this.to - this.from) * t;
        const cos = Math.cos(this.inclination);
        const sin = Math.sin(this.inclination);
        const dx = -this.semiMajor * Math.sin(theta);
        const dy = this.semiMinor * Math.cos(theta);
        const x = dx * cos - dy * sin;
        const y = dx * sin + dy * cos;
        const length = Math.sqrt(x * x + y * y);
        if (!(length > 0)) {
            return new Point(1, 0);
        }

        const sign = this.isCounterclockwise ? 1 : -1;
        return new Point(sign * x / length, sign * y / length);
    }

    get length() {
        return GeometryMath.ellipseArcLength(this.semiMajor, this.semiMinor, Math.min(this.from, this.to), Math.abs(this.to - this.from));
    }

    areaTerm(origin) {
        const start = this.start;
        const end = this.end;
        return (this.semiMajor * this.semiMinor * (this.to - this.from)
            + (this.center.x - origin.x) * (end.y - start.y)
            - (this.center.y - origin.y) * (end.x - start.x)) / 2;
    }

    part(t0, t1) {
        return new ArcPiece(this.center, this.semiMajor, this.semiMinor, this.inclination, this.from + (this.to - this.from) * t0, this.from + (this.to - this.from) * t1);
    }

    reversed() {
        return new ArcPiece(this.center, this.semiMajor, this.semiMinor, this.inclination, this.to, this.from);
    }

    /** The parametric angle of the point of the ellipse on the ray from the center through the point */
    angleOf(point) {
        const cos = Math.cos(this.inclination);
        const sin = Math.sin(this.inclination);
        const dx = point.x - this.center.x;
        const dy = point.y - this.center.y;
        const along = dx * cos + dy * sin;
        const across = -dx * sin + dy * cos;
        return Math.atan2(across / this.semiMinor, along / this.semiMajor);
    }

    /** The left side of the ellipse's equation minus 1 at the point: 0 on the curve, below inside */
    implicit(point) {
        const cos = Math.cos(this.inclination);
        const sin = Math.sin(this.inclination);
        const dx = point.x - this.center.x;
        const dy = point.y - this.center.y;
        const along = (dx * cos + dy * sin) / this.semiMajor;
        const across = (-dx * sin + dy * cos) / this.semiMinor;
        return along * along + across * across - 1;
    }

    /** Where the parametric angle is along the arc, 0 to 1, or NaN when the arc doesn't reach it */
    parameterOf(theta, tolerance) {
        const sweep = Math.abs(this.to - this.from);
        let u = this.isCounterclockwise ? theta - this.from : this.from - theta;
        u = u % GeometryMath.DOUBLEPI;
        if (u < 0) {
            u += GeometryMath.DOUBLEPI;
        }

        if (u > GeometryMath.DOUBLEPI - tolerance) {
            u -= GeometryMath.DOUBLEPI;
        }

        if (u < -tolerance || u > sweep + tolerance) {
            return NaN;
        }

        return sweep > 0 ? Math.max(0, Math.min(1, u / sweep)) : 0;
    }

    nearestPoint(point) {
        const theta = this.angleOf(point);
        if (!Number.isNaN(this.parameterOf(theta, 0))) {
            return this.at(theta);
        }

        const start = this.start;
        const end = this.end;
        return start.distance(point) <= end.distance(point) ? start : end;
    }

    get bounds() {
        const cos = Math.cos(this.inclination);
        const sin = Math.sin(this.inclination);
        const width = Math.sqrt((this.semiMajor * cos) ** 2 + (this.semiMinor * sin) ** 2);
        const height = Math.sqrt((this.semiMajor * sin) ** 2 + (this.semiMinor * cos) ** 2);
        return new Rect(this.center.x - width, this.center.y - height, 2 * width, 2 * height);
    }

    addSamples(points) {
        const count = Math.max(2, Math.ceil(Math.abs(this.to - this.from) / (Math.PI / 24)));
        for (let i = 1; i <= count; i++) {
            points.push(this.pointAt(i / count));
        }
    }
}

class RegionBoundary {
    constructor() {
        /** The pieces of the boundary, the result on the left of each */
        this.pieces = [];

        /** The pieces chained into closed loops, for drawing (even-odd) */
        this.loops = [];
        this.area = 0;
        this.perimeter = 0;
        this.centroid = Point.infinite;
        this.scale = 1;
    }

    static get empty() {
        return new RegionBoundary();
    }

    get hasInside() {
        return this.pieces.length > 0 && this.area > 1e-12 * this.scale * this.scale;
    }

    nearestPoint(point) {
        let result = Point.infinite;
        let best = Infinity;
        for (const piece of this.pieces) {
            const candidate = piece.nearestPoint(point);
            const distance = candidate.distance(point);
            if (candidate.exists() && distance < best) {
                best = distance;
                result = candidate;
            }
        }

        return result;
    }

    distanceTo(point) {
        const nearest = this.nearestPoint(point);
        return nearest.exists() ? nearest.distance(point) : Infinity;
    }

    /** Whether the operation holds for the memberships, one per operand in order */
    static evaluate(operation, memberships) {
        if (memberships.length === 0) {
            return false;
        }

        switch (operation) {
            case RegionOperation.Difference:
                if (!memberships[0]) {
                    return false;
                }

                for (let i = 1; i < memberships.length; i++) {
                    if (memberships[i]) {
                        return false;
                    }
                }

                return true;
            case RegionOperation.Intersection:
                return memberships.every(m => m);
            default:
                return memberships.some(m => m);
        }
    }

    static contains(operation, operands, point) {
        const memberships = operands.map(operand => operand != null && operand.exists && operand.hasInside && operand.containsPoint(point));
        return RegionBoundary.evaluate(operation, memberships);
    }

    static near(a, b, tolerance) {
        return Math.abs(a.x - b.x) <= tolerance && Math.abs(a.y - b.y) <= tolerance;
    }

    // Tracing

    static trace(operation, operands) {
        const result = new RegionBoundary();
        if (operands.length === 0 || operands.some(operand => operand == null || !operand.exists)) {
            return result;
        }

        let pieces = [];
        for (let i = 0; i < operands.length; i++) {
            if (operands[i].hasInside) {
                for (const piece of RegionBoundary.piecesOf(operands[i])) {
                    pieces.push({ piece, operand: i });
                }
            }
        }

        if (pieces.length === 0) {
            return result;
        }

        const scale = RegionBoundary.scaleOf(pieces.map(p => p.piece));
        const tolerance = RegionBoundary.tolerance(scale);
        result.scale = scale;
        pieces = RegionBoundary.split(pieces, tolerance);

        const oriented = [];
        for (const { piece, operand } of pieces) {
            const turned = RegionBoundary.orientToOperand(piece, operands[operand], scale);
            if (turned != null) {
                oriented.push({ piece: turned, operand });
            }
        }

        const kept = [];
        const memberships = new Array(operands.length);
        for (const { piece, operand } of oriented) {
            const middle = piece.middle;
            for (let j = 0; j < operands.length; j++) {
                memberships[j] = j !== operand && RegionBoundary.isInside(operands[j], middle);
            }

            for (const other of oriented) {
                if (other.operand !== operand && other.piece.isSameAs(piece, tolerance)) {
                    memberships[other.operand] = RegionBoundary.near(other.piece.start, piece.start, tolerance);
                }
            }

            memberships[operand] = true;
            const inLeft = RegionBoundary.evaluate(operation, memberships);
            for (const other of oriented) {
                if (other.operand !== operand && other.piece.isSameAs(piece, tolerance)) {
                    memberships[other.operand] = !RegionBoundary.near(other.piece.start, piece.start, tolerance);
                }
            }

            memberships[operand] = false;
            const inRight = RegionBoundary.evaluate(operation, memberships);
            if (inLeft === inRight) {
                continue;
            }

            const boundary = inLeft ? piece : piece.reversed();
            if (!kept.some(k => k.isSameAs(boundary, tolerance))) {
                kept.push(boundary);
            }
        }

        result.pieces = kept;
        result.loops = RegionBoundary.chain(kept, tolerance);
        result.measure();
        return result;
    }

    static tolerance(scale) {
        return 1e-9 * scale;
    }

    static scaleOf(pieces) {
        let result = 1;
        for (const piece of pieces) {
            const bounds = piece.bounds;
            result = Math.max(result, Math.abs(bounds.x), Math.abs(bounds.right), Math.abs(bounds.y), Math.abs(bounds.bottom));
        }

        return result;
    }

    static isInside(operand, point) {
        return operand.hasInside && operand.containsPoint(point);
    }

    /** The piece with its operand on the left, or null when the operand is on neither side or on both */
    static orientToOperand(piece, operand, scale) {
        const length = piece.length;
        const step = Math.max(1e-13 * scale, Math.min(1e-7 * scale, length / 100));
        const middle = piece.middle;
        const normal = piece.leftNormalAt(0.5);
        const left = RegionBoundary.isInside(operand, new Point(middle.x + normal.x * step, middle.y + normal.y * step));
        const right = RegionBoundary.isInside(operand, new Point(middle.x - normal.x * step, middle.y - normal.y * step));
        if (left === right) {
            return null;
        }

        return left ? piece : piece.reversed();
    }

    // The operands' boundaries

    static get maxArcSweep() {
        return Math.PI / 2;
    }

    /** The boundary of a figure with an inside, in pieces, in any order and either way round */
    static piecesOf(figure) {
        const result = [];
        if (figure.isRegionFigure === true) {
            result.push(...figure.boundary.pieces);
        } else if (figure.isPolygonalChain === true) {
            RegionBoundary.addPolygon(result, figure.currentVertexCoordinates());
        } else if (figure.isArc === true && figure.isShapeWithInterior === true) {
            const [begin, sweep] = figure.parametricArc;
            RegionBoundary.addArc(result, figure.center, figure.semiMajor, figure.semiMinor, figure.inclination, begin, begin + sweep);
            const start = figure.beginLocation;
            const end = figure.endLocation;
            if (figure.isSectorShape) {
                const center = figure.center;
                result.push(new SegmentPiece(end, center));
                result.push(new SegmentPiece(center, start));
            } else {
                result.push(new SegmentPiece(end, start));
            }
        } else if (figure.isEllipse === true) {
            RegionBoundary.addArc(result, figure.center, figure.semiMajor, figure.semiMinor, figure.inclination, 0, GeometryMath.DOUBLEPI);
        } else if (figure.isBezierPath === true) {
            RegionBoundary.addPolygon(result, figure.outlinePolygon());
            for (const hole of figure.holes) {
                if (hole.exists && hole.pieceCount > 0) {
                    RegionBoundary.addPolygon(result, hole.outlinePolygon());
                }
            }
        }

        return result.filter(piece => piece.length > 0 && piece.start.exists() && piece.end.exists());
    }

    static addPolygon(result, vertices) {
        const count = vertices.length;
        if (count < 3) {
            return;
        }

        for (let i = 0; i < count; i++) {
            const from = vertices[i];
            const to = vertices[(i + 1) % count];
            if (!from.equals(to)) {
                result.push(new SegmentPiece(from, to));
            }
        }
    }

    /** An arc in pieces of at most a quarter turn each */
    static addArc(result, center, semiMajor, semiMinor, inclination, from, to) {
        if (!(semiMajor > 0 && semiMinor > 0) || !(Math.abs(to - from) > 0)) {
            return;
        }

        const count = Math.max(1, Math.ceil(Math.abs(to - from) / RegionBoundary.maxArcSweep - 1e-9));
        for (let i = 0; i < count; i++) {
            const a = from + (to - from) * i / count;
            const b = from + (to - from) * (i + 1) / count;
            result.push(new ArcPiece(center, semiMajor, semiMinor, inclination, a, b));
        }
    }

    // Cutting at the crossings

    static split(pieces, tolerance) {
        const cuts = pieces.map(() => []);
        const bounds = pieces.map(p => RegionBoundary.inflate(p.piece.bounds, tolerance));
        for (let i = 0; i < pieces.length; i++) {
            for (let j = i + 1; j < pieces.length; j++) {
                if (RegionBoundary.intersects(bounds[i], bounds[j])) {
                    RegionBoundary.findCrossings(pieces[i].piece, pieces[j].piece, cuts[i], cuts[j], tolerance);
                }
            }
        }

        const result = [];
        for (let i = 0; i < pieces.length; i++) {
            const piece = pieces[i].piece;
            const parameterTolerance = tolerance / piece.length;
            const parameters = cuts[i].filter(t => t > parameterTolerance && t < 1 - parameterTolerance).sort((a, b) => a - b);
            let last = 0;
            for (const t of parameters) {
                if (t - last > parameterTolerance) {
                    result.push({ piece: piece.part(last, t), operand: pieces[i].operand });
                    last = t;
                }
            }

            result.push({ piece: last === 0 ? piece : piece.part(last, 1), operand: pieces[i].operand });
        }

        return result;
    }

    static inflate(rect, size) {
        return new Rect(rect.x - size, rect.y - size, rect.width + 2 * size, rect.height + 2 * size);
    }

    static intersects(a, b) {
        return a.x <= b.right && b.x <= a.right && a.y <= b.bottom && b.y <= a.bottom;
    }

    static findCrossings(p, q, cutsP, cutsQ, tolerance) {
        if (p instanceof SegmentPiece && q instanceof SegmentPiece) {
            RegionBoundary.crossSegments(p, q, cutsP, cutsQ, tolerance);
        } else if (p instanceof SegmentPiece) {
            RegionBoundary.crossSegmentAndArc(p, q, cutsP, cutsQ, tolerance);
        } else if (q instanceof SegmentPiece) {
            RegionBoundary.crossSegmentAndArc(q, p, cutsQ, cutsP, tolerance);
        } else {
            RegionBoundary.crossArcs(p, q, cutsP, cutsQ, tolerance);
        }
    }

    static cross(a, b) {
        return a.x * b.y - a.y * b.x;
    }

    static crossSegments(p, q, cutsP, cutsQ, tolerance) {
        const r = p.end.minus(p.start);
        const s = q.end.minus(q.start);
        const offset = q.start.minus(p.start);
        const lengthP = p.length;
        const lengthQ = q.length;
        const denominator = RegionBoundary.cross(r, s);
        if (Math.abs(denominator) <= 1e-12 * lengthP * lengthQ) {
            // parallel: on one line, each is cut where the other's ends lie on it
            if (Math.abs(RegionBoundary.cross(offset, r)) <= tolerance * lengthP) {
                RegionBoundary.addWithin(cutsP, p.parameterOf(q.start), tolerance / lengthP);
                RegionBoundary.addWithin(cutsP, p.parameterOf(q.end), tolerance / lengthP);
                RegionBoundary.addWithin(cutsQ, q.parameterOf(p.start), tolerance / lengthQ);
                RegionBoundary.addWithin(cutsQ, q.parameterOf(p.end), tolerance / lengthQ);
            }

            return;
        }

        const t = RegionBoundary.cross(offset, s) / denominator;
        const u = RegionBoundary.cross(offset, r) / denominator;
        if (RegionBoundary.isWithin(t, tolerance / lengthP) && RegionBoundary.isWithin(u, tolerance / lengthQ)) {
            cutsP.push(t);
            cutsQ.push(u);
        }
    }

    static isWithin(t, parameterTolerance) {
        return t >= -parameterTolerance && t <= 1 + parameterTolerance;
    }

    static addWithin(cuts, t, parameterTolerance) {
        if (RegionBoundary.isWithin(t, parameterTolerance)) {
            cuts.push(t);
        }
    }

    static crossSegmentAndArc(segment, arc, cutsSegment, cutsArc, tolerance) {
        const crossings = GeometryMath.getIntersectionOfEllipseAndLine(
            arc.center,
            arc.semiMajor,
            arc.semiMinor,
            arc.inclination,
            new PointPair(segment.start, segment.end));
        RegionBoundary.addCrossing(segment, arc, crossings.p1, cutsSegment, cutsArc, tolerance);
        if (!RegionBoundary.near(crossings.p1, crossings.p2, tolerance)) {
            RegionBoundary.addCrossing(segment, arc, crossings.p2, cutsSegment, cutsArc, tolerance);
        }
    }

    static addCrossing(segment, arc, point, cutsSegment, cutsArc, tolerance) {
        if (!point.exists()) {
            return;
        }

        const t = segment.parameterOf(point);
        const u = arc.parameterOf(arc.angleOf(point), RegionBoundary.angularTolerance(arc, tolerance));
        if (RegionBoundary.isWithin(t, tolerance / segment.length) && !Number.isNaN(u)) {
            cutsSegment.push(t);
            cutsArc.push(u);
        }
    }

    static angularTolerance(arc, tolerance) {
        return tolerance / Math.max(arc.semiMajor, arc.semiMinor);
    }

    static crossArcs(p, q, cutsP, cutsQ, tolerance) {
        const sameEllipse = RegionBoundary.near(p.center, q.center, tolerance)
            && Math.abs(p.semiMajor - q.semiMajor) <= tolerance
            && Math.abs(p.semiMinor - q.semiMinor) <= tolerance
            && (p.isCircle && q.isCircle || Math.abs(GeometryMath.ieeeRemainder(p.inclination - q.inclination, Math.PI)) <= 1e-9);
        if (sameEllipse) {
            RegionBoundary.addArcCut(cutsP, p, q.start, tolerance);
            RegionBoundary.addArcCut(cutsP, p, q.end, tolerance);
            RegionBoundary.addArcCut(cutsQ, q, p.start, tolerance);
            RegionBoundary.addArcCut(cutsQ, q, p.end, tolerance);
            return;
        }

        if (p.isCircle && q.isCircle) {
            const crossings = GeometryMath.getIntersectionOfCircles(p.center, p.semiMajor, q.center, q.semiMajor);
            RegionBoundary.addArcCrossing(p, q, crossings.p1, cutsP, cutsQ, tolerance);
            if (!RegionBoundary.near(crossings.p1, crossings.p2, tolerance)) {
                RegionBoundary.addArcCrossing(p, q, crossings.p2, cutsP, cutsQ, tolerance);
            }

            return;
        }

        // an ellipse among them: the roots of its equation along the other arc, by sampling and bisection
        const steps = 48;
        let previous = p.implicit(q.pointAt(0));
        for (let i = 1; i <= steps; i++) {
            const t = i / steps;
            const value = p.implicit(q.pointAt(t));
            if (previous === 0 || previous * value < 0) {
                let a = (i - 1) / steps;
                let b = t;
                let valueA = previous;
                if (previous !== 0) {
                    for (let k = 0; k < 60; k++) {
                        const m = (a + b) / 2;
                        const valueM = p.implicit(q.pointAt(m));
                        if (valueM === 0) {
                            a = b = m;
                            break;
                        }

                        if (valueA * valueM < 0) {
                            b = m;
                        } else {
                            a = m;
                            valueA = valueM;
                        }
                    }
                } else {
                    b = a;
                }

                RegionBoundary.addArcCrossing(p, q, q.pointAt((a + b) / 2), cutsP, cutsQ, tolerance);
            }

            previous = value;
        }
    }

    static addArcCut(cuts, arc, point, tolerance) {
        const t = arc.parameterOf(arc.angleOf(point), RegionBoundary.angularTolerance(arc, tolerance));
        if (!Number.isNaN(t)) {
            cuts.push(t);
        }
    }

    static addArcCrossing(p, q, point, cutsP, cutsQ, tolerance) {
        if (!point.exists()) {
            return;
        }

        const t = p.parameterOf(p.angleOf(point), RegionBoundary.angularTolerance(p, tolerance));
        const u = q.parameterOf(q.angleOf(point), RegionBoundary.angularTolerance(q, tolerance));
        if (!Number.isNaN(t) && !Number.isNaN(u)) {
            cutsP.push(t);
            cutsQ.push(u);
        }
    }

    // Loops and measures

    /** The pieces chained end to start into closed loops; any pairing where several meet will do (even-odd goes by the pieces) */
    static chain(pieces, tolerance) {
        const result = [];
        const used = new Array(pieces.length).fill(false);
        for (let start = 0; start < pieces.length; start++) {
            if (used[start]) {
                continue;
            }

            const loop = [pieces[start]];
            used[start] = true;
            const origin = pieces[start].start;
            let end = pieces[start].end;
            while (!RegionBoundary.near(end, origin, tolerance)) {
                let next = -1;
                let best = Infinity;
                for (let i = 0; i < pieces.length; i++) {
                    if (used[i]) {
                        continue;
                    }

                    const distance = pieces[i].start.distance(end);
                    if (distance <= tolerance && distance < best) {
                        best = distance;
                        next = i;
                    }
                }

                if (next < 0) {
                    break;
                }

                used[next] = true;
                loop.push(pieces[next]);
                end = pieces[next].end;
            }

            result.push(loop);
        }

        return result;
    }

    measure() {
        if (this.pieces.length === 0) {
            this.area = 0;
            this.perimeter = 0;
            this.centroid = Point.infinite;
            return;
        }

        const origin = this.pieces[0].start;
        let area = 0;
        let perimeter = 0;
        for (const piece of this.pieces) {
            area += piece.areaTerm(origin);
            perimeter += piece.length;
        }

        this.area = area;
        this.perimeter = perimeter;
        this.centroid = this.findCentroid(origin);
    }

    findCentroid(origin) {
        let area = 0;
        let x = 0;
        let y = 0;
        for (const loop of this.loops) {
            const points = [loop[0].start.minus(origin)];
            for (const piece of loop) {
                const samples = [];
                piece.addSamples(samples);
                points.push(...samples.map(sample => sample.minus(origin)));
            }

            for (let i = 0; i < points.length; i++) {
                const a = points[i];
                const b = points[(i + 1) % points.length];
                const cross = a.x * b.y - b.x * a.y;
                area += cross;
                x += (a.x + b.x) * cross;
                y += (a.y + b.y) * cross;
            }
        }

        if (Math.abs(area) > 1e-12 * this.scale * this.scale) {
            return new Point(origin.x + x / (3 * area), origin.y + y / (3 * area));
        }

        const ends = this.pieces.map(piece => piece.start);
        return new Point(ends.reduce((sum, p) => sum + p.x, 0) / ends.length, ends.reduce((sum, p) => sum + p.y, 0) / ends.length);
    }

    // Drawing

    /** The loops as path commands in pixels (CanvasRenderer.drawPath, even-odd), and the box they lie in */
    buildPath(coordinateSystem) {
        const commands = [];
        const points = [];
        for (const loop of this.loops) {
            const start = coordinateSystem.toPhysical(loop[0].start);
            commands.push({ op: "move", x: start.x, y: start.y });
            points.push(start);
            for (const piece of loop) {
                const end = coordinateSystem.toPhysical(piece.end);
                if (piece instanceof ArcPiece) {
                    const center = coordinateSystem.toPhysical(piece.center);
                    // canvas angles run the other way (y down): a counterclockwise sweep here
                    // is one of decreasing canvas angle (ArcBase.buildPath)
                    commands.push({
                        op: "arc",
                        cx: center.x,
                        cy: center.y,
                        rx: coordinateSystem.toPhysicalLength(piece.semiMajor),
                        ry: coordinateSystem.toPhysicalLength(piece.semiMinor),
                        rotation: -piece.inclination,
                        start: -piece.from,
                        end: -piece.to,
                        counterclockwise: piece.isCounterclockwise
                    });
                    const samples = [];
                    piece.addSamples(samples);
                    points.push(...samples.map(sample => coordinateSystem.toPhysical(sample)));
                } else {
                    commands.push({ op: "line", x: end.x, y: end.y });
                    points.push(end);
                }
            }

            commands.push({ op: "close" });
        }

        return { commands, box: points.length > 0 ? CanvasRenderer.boxOf(points) : null };
    }
}
