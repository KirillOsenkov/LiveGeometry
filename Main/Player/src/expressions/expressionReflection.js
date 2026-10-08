// Port of Main/Avalonia/DynamicGeometry/Expressions/ExpressionReflection.cs: what binding
// asks reflection for in C# - a function by its name and the number of its arguments, a
// property of a figure by name - as tables here. A function is { name, parameterCount,
// takesNumbers, takesPoints, takesArray, invoke }.

const ExpressionReflection = {
    /** The functions of System.Math by name (as the C# finds them), with how many doubles each takes */
    mathFunctions: [
        ["Sin", 1, Math.sin], ["Cos", 1, Math.cos], ["Tan", 1, Math.tan],
        ["Asin", 1, Math.asin], ["Acos", 1, Math.acos], ["Atan", 1, Math.atan], ["Atan2", 2, Math.atan2],
        ["Sinh", 1, Math.sinh], ["Cosh", 1, Math.cosh], ["Tanh", 1, Math.tanh],
        ["Asinh", 1, Math.asinh], ["Acosh", 1, Math.acosh], ["Atanh", 1, Math.atanh],
        ["Sqrt", 1, Math.sqrt], ["Cbrt", 1, Math.cbrt], ["Abs", 1, Math.abs], ["Exp", 1, Math.exp],
        ["Log", 1, Math.log], ["Log", 2, (a, newBase) => Math.log(a) / Math.log(newBase)],
        ["Log10", 1, Math.log10], ["Log2", 1, Math.log2],
        ["Pow", 2, Math.pow],
        ["Floor", 1, Math.floor], ["Ceiling", 1, Math.ceil], ["Truncate", 1, Math.trunc],
        ["Round", 1, x => Number.isNaN(x) ? NaN : (Math.abs(x % 1) === 0.5 ? 2 * Math.round(x / 2) : Math.round(x))],
        ["Min", 2, Math.min], ["Max", 2, Math.max],
        ["Sign", 1, Math.sign],
        ["Clamp", 3, (value, min, max) => Math.max(min, Math.min(max, value))],
        ["IEEERemainder", 2, (x, y) => GeometryMath.ieeeRemainder(x, y)],
        ["CopySign", 2, (x, y) => Math.sign(y) === 0 && !Object.is(y, -0) ? Math.abs(x) : Math.abs(x) * (Object.is(y, -0) || y < 0 ? -1 : 1)],
        ["FusedMultiplyAdd", 3, (x, y, z) => x * y + z],
        ["ScaleB", 2, (x, n) => x * Math.pow(2, n)],
        ["BitIncrement", 1, x => x + Number.EPSILON * Math.abs(x)],
        ["BitDecrement", 1, x => x - Number.EPSILON * Math.abs(x)],
        ["MaxMagnitude", 2, (x, y) => Math.abs(x) >= Math.abs(y) ? x : y],
        ["MinMagnitude", 2, (x, y) => Math.abs(x) <= Math.abs(y) ? x : y],
        ["ILogB", 1, x => Math.floor(Math.log2(Math.abs(x)))],
        ["ReciprocalEstimate", 1, x => 1 / x],
        ["ReciprocalSqrtEstimate", 1, x => 1 / Math.sqrt(x)]
    ],

    methods: null,
    resolvedMethods: new Map(),

    /** The methods in the order the C# searches them: ours first, then System.Math's */
    allMethods() {
        if (ExpressionReflection.methods == null) {
            const methods = [];
            for (const name of Object.keys(Functions)) {
                const pointCount = PointFunctionNames[name];
                if (pointCount === undefined) {
                    const invoke = Functions[name];
                    methods.push({ name, parameterCount: invoke.length, takesNumbers: true, takesPoints: false, takesArray: false, invoke });
                } else {
                    methods.push({
                        name,
                        parameterCount: pointCount < 0 ? 1 : pointCount,
                        takesNumbers: false,
                        takesPoints: true,
                        takesArray: pointCount < 0,
                        invoke: Functions[name]
                    });
                }
            }

            for (const [name, parameterCount, invoke] of ExpressionReflection.mathFunctions) {
                methods.push({ name, parameterCount, takesNumbers: true, takesPoints: false, takesArray: false, invoke });
            }

            ExpressionReflection.methods = methods;
        }

        return ExpressionReflection.methods;
    },

    /**
     * The function called by this name with this many arguments: one that takes that many
     * numbers if there is one - ours first, then System.Math's - else whatever goes by the
     * name (ours that take points). Null when there is none.
     */
    resolveMethod(functionName, argumentCount) {
        const key = functionName.toLowerCase() + "/" + argumentCount;
        if (!ExpressionReflection.resolvedMethods.has(key)) {
            ExpressionReflection.resolvedMethods.set(key, ExpressionReflection.findMethod(functionName, argumentCount));
        }

        return ExpressionReflection.resolvedMethods.get(key);
    },

    findMethod(functionName, argumentCount) {
        const lower = functionName.toLowerCase();
        for (const method of ExpressionReflection.allMethods()) {
            if (method.name.toLowerCase() === lower && ExpressionReflection.takesNumbers(method, argumentCount)) {
                return method;
            }
        }

        for (const method of ExpressionReflection.allMethods()) {
            if (method.name.toLowerCase() === lower) {
                return method;
            }
        }

        return null;
    },

    /** A function of numbers, called with as many as it takes: its arguments are expressions, not the names of points */
    takesNumbers(method, argumentCount) {
        return method.parameterCount === argumentCount && argumentCount > 0 && method.takesNumbers;
    },

    /**
     * The property a name means on a figure: any public numeric property, by name in any
     * case, the one written exactly first (A.X is x, AB.Length is length). Null when the
     * figure has none, or it is no number.
     */
    findProperty(figure, propertyName) {
        const names = ExpressionReflection.propertyNames(figure);
        const lower = propertyName.toLowerCase();
        const candidates = names.filter(name => name.toLowerCase() === lower);
        const chosen = candidates.find(name => name === propertyName) ?? candidates[0] ?? null;
        return chosen;
    },

    propertyNamesByType: new Map(),

    propertyNames(figure) {
        const type = figure.constructor;
        if (!ExpressionReflection.propertyNamesByType.has(type)) {
            // the fields set in the constructors (coordinates, parameter) and the getters
            // of the class and its bases (x, length, radius)
            const names = new Set(Object.keys(figure));
            let prototype = Object.getPrototypeOf(figure);
            while (prototype != null && prototype !== Object.prototype) {
                for (const name of Object.getOwnPropertyNames(prototype)) {
                    const descriptor = Object.getOwnPropertyDescriptor(prototype, name);
                    if (descriptor != null && descriptor.get != null && name !== "constructor") {
                        names.add(name);
                    }
                }

                prototype = Object.getPrototypeOf(prototype);
            }

            ExpressionReflection.propertyNamesByType.set(type, [...names]);
        }

        return ExpressionReflection.propertyNamesByType.get(type);
    },

    isNumberValue(value) {
        return typeof value === "number";
    }
};
