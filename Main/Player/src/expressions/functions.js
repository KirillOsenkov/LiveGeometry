// Port of Main/Avalonia/LiveGeometry/Expressions/Functions.cs and FunctionCatalog.cs: the
// functions of the expression language, all of them. Trigonometry works in radians; what
// is read off the drawing is in degrees. Each is a function of numbers, or of points (its
// parameters are Points; area takes any number of them).

const Functions = {
    Sin(x) {
        return Math.sin(x);
    },

    Cos(x) {
        return Math.cos(x);
    },

    Tan(x) {
        return Math.tan(x);
    },

    Sec(x) {
        return 1 / Math.cos(x);
    },

    Csc(x) {
        return 1 / Math.sin(x);
    },

    Cot(x) {
        return Math.cos(x) / Math.sin(x);
    },

    Arcsin(x) {
        return Math.asin(x);
    },

    Arccos(x) {
        return Math.acos(x);
    },

    Arctan(x) {
        return Math.atan(x);
    },

    Atan2(y, x) {
        return Math.atan2(y, x);
    },

    Sinh(x) {
        return Math.sinh(x);
    },

    Cosh(x) {
        return Math.cosh(x);
    },

    Tanh(x) {
        return Math.tanh(x);
    },

    Arcsinh(x) {
        return Math.asinh(x);
    },

    Arccosh(x) {
        return Math.acosh(x);
    },

    Arctanh(x) {
        return Math.atanh(x);
    },

    Sqrt(x) {
        return Math.sqrt(x);
    },

    Cbrt(x) {
        return Math.cbrt(x);
    },

    /** The nth root; of a negative number only for an odd whole n */
    Nthroot(x, n) {
        if (x < 0 && n === Math.floor(n) && Math.abs(n % 2) === 1) {
            return -Math.pow(-x, 1 / n);
        }

        return Math.pow(x, 1 / n);
    },

    Exp(x) {
        return Math.exp(x);
    },

    Ln(x) {
        return Math.log(x);
    },

    /** The common logarithm, to base 10 */
    Log(x) {
        return Math.log10(x);
    },

    /** log_2(x): the base is in the function's name */
    LogBase(x, logBase) {
        return Math.log(x) / Math.log(logBase);
    },

    Abs(x) {
        return Math.abs(x);
    },

    /** A half goes up, as at school; round(x, digits) to so many digits after the point */
    Round(x, digits) {
        if (digits === undefined) {
            return Number.isNaN(x) ? NaN : Math.sign(x) * Math.round(Math.abs(x));
        }

        if (digits !== Math.floor(digits) || !Number.isFinite(x)) {
            return NaN;
        }

        if (digits < 0) {
            const factor = Math.pow(10, -digits);
            return Functions.Round(x / factor) * factor;
        }

        return digits > 15 ? x : GeometryMath.round(x, digits);
    },

    Floor(x) {
        return Math.floor(x);
    },

    Ceil(x) {
        return Math.ceil(x);
    },

    Sign(x) {
        return Number.isNaN(x) ? NaN : Math.sign(x);
    },

    /** What is left of a over whole b's: never negative for a b above 0 */
    Mod(a, b) {
        return b === 0 ? NaN : a - b * Math.floor(a / b);
    },

    Min(a, b) {
        return Math.min(a, b);
    },

    Max(a, b) {
        return Math.max(a, b);
    },

    Clamp(x, low, high) {
        return low > high ? NaN : Math.max(low, Math.min(high, x));
    },

    /** The greatest common factor of two whole numbers */
    Gcd(a, b) {
        const isWhole = value => Number.isFinite(value) && value === Math.floor(value) && Math.abs(value) < 1e15;
        if (!isWhole(a) || !isWhole(b)) {
            return NaN;
        }

        a = Math.abs(a);
        b = Math.abs(b);
        while (b !== 0) {
            [a, b] = [b, a % b];
        }

        return a;
    },

    Lcm(a, b) {
        const gcd = Functions.Gcd(a, b);
        if (Number.isNaN(gcd)) {
            return NaN;
        }

        return gcd === 0 ? 0 : Math.abs(a / gcd * b);
    },

    Deg(x) {
        return GeometryMath.toDegrees(x);
    },

    Rad(x) {
        return GeometryMath.toRadians(x);
    },

    /** a where the condition holds, b where it doesn't (no value without b) */
    If(condition, a, b = NaN) {
        if (Number.isNaN(condition)) {
            return NaN;
        }

        return condition !== 0 ? a : b;
    },

    Less(a, b) {
        return Functions.compare(a, b, (x, y) => x < y && !Functions.areEqual(x, y));
    },

    LessOrEqual(a, b) {
        return Functions.compare(a, b, (x, y) => x < y || Functions.areEqual(x, y));
    },

    Greater(a, b) {
        return Functions.compare(a, b, (x, y) => x > y && !Functions.areEqual(x, y));
    },

    GreaterOrEqual(a, b) {
        return Functions.compare(a, b, (x, y) => x > y || Functions.areEqual(x, y));
    },

    Equal(a, b) {
        return Functions.compare(a, b, Functions.areEqual);
    },

    NotEqual(a, b) {
        return Functions.compare(a, b, (x, y) => !Functions.areEqual(x, y));
    },

    Distance(a, b) {
        return a.distance(b);
    },

    /** The smaller angle at b, in degrees, as the Angle tool measures it */
    Angle(a, b, c) {
        const a1 = GeometryMath.getAngle(b, a);
        const a2 = GeometryMath.getAngle(b, c);
        let result = Math.abs(a2 - a1);
        if (result >= Math.PI) {
            result = 2 * Math.PI - result;
        }

        return GeometryMath.toDegrees(result);
    },

    /** The angle at b counterclockwise from BA to BC, in degrees, from 0 up to 360 */
    Directedangle(a, b, c) {
        return GeometryMath.toDegrees(GeometryMath.oAngle(a, b, c));
    },

    /** The direction from a to b in degrees, as a segment's Direction says it */
    Direction(a, b) {
        const angle = GeometryMath.toDegrees(GeometryMath.getAngle(a, b));
        return 360 - angle < 1e-9 ? 0 : angle;
    },

    /** Takes the points themselves, as many as there are */
    Area(points) {
        return GeometryMath.area(points.map(point => point.coordinates));
    }
};

// (helpers, not functions of the language: lowercase, which the parity check leaves out)
Functions.compare = (a, b, holds) => Number.isNaN(a) || Number.isNaN(b) ? NaN : (holds(a, b) ? 1 : 0);
Functions.areEqual = (a, b) => a === b || Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));

/** Which of ours take points (the rest take numbers), and Area takes the points themselves */
const PointFunctionNames = {
    Distance: 2, Angle: 3, Directedangle: 3, Direction: 2, Area: -1
};

/**
 * The functions as FunctionCatalog.cs lists them: the name, the method's name, how many
 * numbers it takes ("variadic": two at a time, any number of them), its other names.
 */
const FunctionCatalog = {
    all: [
        { name: "sin" }, { name: "cos" }, { name: "tan" }, { name: "sec" }, { name: "csc" }, { name: "cot" },
        { name: "arcsin", aliases: ["asin"] }, { name: "arccos", aliases: ["acos"] }, { name: "arctan", aliases: ["atan"] },
        { name: "atan2" }, { name: "sinh" }, { name: "cosh" }, { name: "tanh" },
        { name: "arcsinh", aliases: ["asinh"] }, { name: "arccosh", aliases: ["acosh"] }, { name: "arctanh", aliases: ["atanh"] },
        { name: "sqrt" }, { name: "cbrt" }, { name: "nthroot" }, { name: "exp" }, { name: "ln" }, { name: "log" },
        { name: "abs" }, { name: "round", counts: [1, 2] }, { name: "floor" }, { name: "ceil", aliases: ["ceiling"] }, { name: "sign" },
        { name: "mod" }, { name: "min", variadic: true }, { name: "max", variadic: true }, { name: "clamp" },
        { name: "gcd", variadic: true }, { name: "lcm", variadic: true }, { name: "deg" }, { name: "rad" },
        { name: "if", counts: [2, 3] },
        { name: "distance" }, { name: "angle" }, { name: "directedangle" }, { name: "direction" }, { name: "area" }
    ],

    byName: null,

    /** The function a name calls, by its name or another, in any case; null when the language has none */
    find(name) {
        if (FunctionCatalog.byName == null) {
            FunctionCatalog.byName = new Map();
            for (const function_ of FunctionCatalog.all) {
                for (const each of [function_.name, ...(function_.aliases ?? [])]) {
                    FunctionCatalog.byName.set(each.toLowerCase(), function_);
                }
            }
        }

        return FunctionCatalog.byName.get(name.toLowerCase()) ?? null;
    },

    takesRadians(function_) {
        return function_ != null && ["sin", "cos", "tan", "sec", "csc", "cot"].includes(function_.name);
    },

    givesDegrees(function_) {
        return function_ != null && ["angle", "directedangle", "direction"].includes(function_.name);
    }
};
