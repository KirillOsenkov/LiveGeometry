// Port of Main/Avalonia/DynamicGeometry/Expressions/Functions.cs: the library's own functions
// of the expression language. Angles are in radians unless noted. Each is a function of
// numbers, or of points (its parameters are Points; area takes any number of them).

const Functions = {
    Distance(a, b) {
        return a.distance(b);
    },

    Dist(a, b) {
        return a.distance(b);
    },

    Ang(a, b, c) {
        return Functions.Angle(a, b, c);
    },

    /** The square root (not number.SquareRoot(), which takes the root of the absolute value) */
    Sqr(number) {
        return Math.sqrt(number);
    },

    /** A half goes up, as at school: System.Math.Round takes 2.5 to 2 */
    Round(number) {
        return Number.isNaN(number) ? NaN : Math.sign(number) * Math.round(Math.abs(number));
    },

    /** System.Math.Sign throws for what is not a number */
    Sign(number) {
        return Number.isNaN(number) ? NaN : Math.sign(number);
    },

    /** System.Math.Clamp throws when the bounds are the wrong way round */
    Clamp(number, low, high) {
        return low > high ? NaN : Math.max(low, Math.min(high, number));
    },

    Ln(number) {
        return Math.log(number);
    },

    /** The unoriented angle at b, in [0, pi] */
    Angle(a, b, c) {
        const a1 = GeometryMath.getAngle(b, a);
        const a2 = GeometryMath.getAngle(b, c);
        let result = a2 < a1 ? a1 - a2 : a2 - a1;
        if (result >= Math.PI) {
            result = 2 * Math.PI - result;
        }

        return result;
    },

    OAngle(a, b, c) {
        return GeometryMath.oAngle(a, b, c);
    },

    XAngle(a, b) {
        return a.angleTo(b);
    },

    XAng(a, b) {
        return a.angleTo(b);
    },

    Norm(a) {
        return a.length();
    },

    Arg(a) {
        return a.arg();
    },

    /** Takes the points themselves, as many as there are */
    Area(points) {
        return GeometryMath.area(points.map(point => point.coordinates));
    },

    Deg(radians) {
        return GeometryMath.toDegrees(radians);
    },

    Rad(degrees) {
        return GeometryMath.toRadians(degrees);
    },

    OAng(a, b, c) {
        return GeometryMath.oAngle(a, b, c);
    },

    Int(number) {
        return Math.floor(number);
    },

    Sgn(number) {
        return Functions.Sign(number);
    },

    Lg(number) {
        return Math.log10(number);
    },

    ToDeg(radians) {
        return GeometryMath.toDegrees(radians);
    },

    ToRad(degrees) {
        return GeometryMath.toRadians(degrees);
    }
};

/** Which of ours take points (the rest take numbers), and Area takes the points themselves */
const PointFunctionNames = {
    Distance: 2, Dist: 2, Ang: 3, Angle: 3, OAngle: 3, XAngle: 2, XAng: 2, Norm: 1, Arg: 1, OAng: 3, Area: -1
};
