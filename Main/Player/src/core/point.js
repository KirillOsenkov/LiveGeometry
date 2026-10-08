// The geometry's value types: what Avalonia's Point, Size and Rect and the library's
// PointPair (Math.cs) are in C#. A Point is immutable: every operation gives a new one.
// The point methods mirror PointExtensions in Math.cs.

class Point {
    constructor(x = 0, y = 0) {
        this.x = x;
        this.y = y;
    }

    /** No point at all: NaN coordinates (Math.InfinitePoint) */
    static get infinite() {
        return new Point(NaN, NaN);
    }

    withX(x) {
        return new Point(x, this.y);
    }

    withY(y) {
        return new Point(this.x, y);
    }

    equals(other) {
        return other != null && this.x === other.x && this.y === other.y;
    }

    distance(other) {
        const dx = this.x - other.x;
        const dy = this.y - other.y;
        return Math.sqrt(dx * dx + dy * dy);
    }

    /** The counterclockwise angle from the x axis direction to the point, in [0, 2pi) */
    angleTo(point) {
        return GeometryMath.oAngle(this.plus(new Point(10, 0)), this, point);
    }

    reflect(center) {
        return new Point(2 * center.x - this.x, 2 * center.y - this.y);
    }

    length() {
        return Math.sqrt(this.sumOfSquares());
    }

    arg() {
        return new Point().angleTo(this);
    }

    scale(factor) {
        return new Point(this.x * factor, this.y * factor);
    }

    snapToIntegers() {
        return new Point(Math.ceil(this.x), Math.ceil(this.y));
    }

    pointInDirection(direction, vectorLength) {
        const vector = direction.minus(this);
        const factor = vectorLength / vector.length();
        return this.plus(vector.scale(factor));
    }

    trimToMaxLength(maxLength) {
        const length = this.length();
        if (maxLength > 0 && length > maxLength) {
            const ratio = maxLength / length;
            return new Point(this.x * ratio, this.y * ratio);
        }

        return this;
    }

    negate() {
        return new Point(-this.x, -this.y);
    }

    minus(other) {
        if (typeof other === "number") {
            return new Point(this.x - other, this.y - other);
        }

        return new Point(this.x - other.x, this.y - other.y);
    }

    plus(other) {
        if (typeof other === "number") {
            return new Point(this.x + other, this.y + other);
        }

        return new Point(this.x + other.x, this.y + other.y);
    }

    offset(xOffset, yOffset = xOffset) {
        return new Point(this.x + xOffset, this.y + yOffset);
    }

    sumOfSquares() {
        return this.x * this.x + this.y * this.y;
    }

    /** Finite coordinates (an infinite one is as much "no point" as NaN) */
    exists() {
        return isValidValue(this.x) && isValidValue(this.y);
    }

    equalsWithPrecision(other) {
        return equalsWithPrecision(this.x, other.x) && equalsWithPrecision(this.y, other.y);
    }

    isEqual(other) {
        return isWithinEpsilon(this.distance(other));
    }

    /** Noise off a coordinate, on every conversion from pixels (not the rounding for display) */
    roundToEpsilon() {
        return new Point(roundToEpsilon(this.x), roundToEpsilon(this.y));
    }

    toString() {
        return "(" + this.x + ", " + this.y + ")";
    }
}

class Size {
    constructor(width = 0, height = 0) {
        this.width = width;
        this.height = height;
    }

    static get infinity() {
        return new Size(Infinity, Infinity);
    }
}

/** An axis-aligned box from (x, y) to (x + width, y + height), whichever way the y axis goes */
class Rect {
    constructor(x = 0, y = 0, width = 0, height = 0) {
        this.x = x;
        this.y = y;
        this.width = width;
        this.height = height;
    }

    static fromPoints(topLeft, size) {
        return new Rect(topLeft.x, topLeft.y, size.width, size.height);
    }

    get right() {
        return this.x + this.width;
    }

    get bottom() {
        return this.y + this.height;
    }

    get topLeft() {
        return new Point(this.x, this.y);
    }

    get bottomRight() {
        return new Point(this.x + this.width, this.y + this.height);
    }

    get center() {
        return new Point(this.x + this.width / 2, this.y + this.height / 2);
    }

    contains(point) {
        return point.x >= this.x && point.x <= this.x + this.width && point.y >= this.y && point.y <= this.y + this.height;
    }

    union(other) {
        const x = Math.min(this.x, other.x);
        const y = Math.min(this.y, other.y);
        const right = Math.max(this.right, other.right);
        const bottom = Math.max(this.bottom, other.bottom);
        return new Rect(x, y, right - x, bottom - y);
    }

    equals(other) {
        return other != null && this.x === other.x && this.y === other.y && this.width === other.width && this.height === other.height;
    }

    translate(offset) {
        return new Rect(this.x + offset.x, this.y + offset.y, this.width, this.height);
    }
}

/** Two points: a segment, a line through them, a box (Math.cs PointPair) */
class PointPair {
    constructor(p1 = new Point(), p2 = new Point()) {
        this.p1 = p1;
        this.p2 = p2;
    }

    static get infinite() {
        return new PointPair(Point.infinite, Point.infinite);
    }

    contains(point) {
        return point.x >= this.p1.x && point.x <= this.p2.x && point.y >= this.p1.y && point.y <= this.p2.y;
    }

    containsInner(point) {
        return point.x > this.p1.x && point.x < this.p2.x && point.y > this.p1.y && point.y < this.p2.y;
    }

    get reverse() {
        return new PointPair(this.p2, this.p1);
    }

    getBoundingRect() {
        let p1 = this.p1;
        let p2 = this.p2;
        if (p1.x > p2.x) {
            const t = p1.x;
            p1 = p1.withX(p2.x);
            p2 = p2.withX(t);
        }

        if (p1.y > p2.y) {
            const t = p1.y;
            p1 = p1.withY(p2.y);
            p2 = p2.withY(t);
        }

        return new PointPair(p1, p2);
    }

    get length() {
        return this.p1.distance(this.p2);
    }

    get midpoint() {
        return new Point((this.p1.x + this.p2.x) / 2, (this.p1.y + this.p2.y) / 2);
    }

    inflate(size) {
        return new PointPair(this.p1.offset(-size), this.p2.offset(size));
    }

    hasValidValue() {
        return this.p1.exists() || this.p2.exists();
    }

    toString() {
        return "(" + this.p1.x + ";" + this.p1.y + ")-(" + this.p2.x + ";" + this.p2.y + ")";
    }
}

// the double helpers of PointExtensions

function isValidValue(value) {
    return typeof value === "number" && !Number.isNaN(value) && Number.isFinite(value);
}

function isValidPositiveValue(value) {
    return isValidValue(value) && value > 0;
}

function isValidNonNegativeValue(value) {
    return isValidValue(value) && value >= 0;
}

function equalsWithPrecision(value, center) {
    return Math.abs(value - center) < GeometryMath.Precision;
}

function isWithinEpsilonTo(value, center) {
    return Math.abs(value - center) < GeometryMath.Epsilon;
}

function isWithinEpsilon(value) {
    return Math.abs(value) < GeometryMath.Epsilon;
}

function roundToEpsilon(value) {
    return roundToDigits(value, 10);
}

/** System.Math.Round(value, digits): to the nearest, halves to the even digit */
function roundToDigits(value, digits) {
    if (!isValidValue(value)) {
        return value;
    }

    const factor = Math.pow(10, digits);
    const scaled = value * factor;
    const rounded = Math.round(scaled);
    // halves to even, as System.Math.Round does by default
    if (Math.abs(scaled - Math.trunc(scaled)) === 0.5 && rounded % 2 !== 0) {
        return (rounded - Math.sign(scaled)) / factor;
    }

    return rounded / factor;
}
