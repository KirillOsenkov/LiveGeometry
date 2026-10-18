// Port of Main/Avalonia/LiveGeometry/Expressions/ExpressionReflection.cs: what binding
// asks reflection for in C# - a function by its name and the number of its arguments, a
// property of a figure by name - as tables here. A function is { name, parameterCount,
// takesNumbers, takesPoints, takesArray, invoke }.

const ExpressionReflection = {
    resolvedMethods: new Map(),

    /**
     * The method a function of the language (FunctionCatalog, by any of its names) is called
     * through with this many arguments: { name, parameterCount, takesNumbers, takesPoints,
     * takesArray, invoke } - for any number of them, the one of two (max(a, b, c) takes them
     * two at a time). Null when the language has no such function, or not with so many.
     */
    resolveMethod(functionName, argumentCount) {
        const function_ = FunctionCatalog.find(functionName);
        if (function_ == null) {
            return null;
        }

        const key = function_.name + "/" + argumentCount;
        if (!ExpressionReflection.resolvedMethods.has(key)) {
            ExpressionReflection.resolvedMethods.set(key, ExpressionReflection.findMethod(function_, argumentCount));
        }

        return ExpressionReflection.resolvedMethods.get(key);
    },

    findMethod(function_, argumentCount) {
        const name = function_.name[0].toUpperCase() + function_.name.substring(1);
        const invoke = Functions[name];
        const pointCount = PointFunctionNames[name];
        if (pointCount !== undefined) {
            const takesArray = pointCount < 0;
            return (takesArray ? argumentCount >= 3 : argumentCount === pointCount)
                ? { name, parameterCount: takesArray ? 1 : pointCount, takesNumbers: false, takesPoints: true, takesArray, invoke }
                : null;
        }

        const counts = function_.counts ?? [invoke.length];
        if (counts.includes(argumentCount)) {
            return { name, parameterCount: argumentCount, takesNumbers: true, takesPoints: false, takesArray: false, invoke };
        }

        if (function_.variadic && argumentCount >= 2) {
            return { name, parameterCount: 2, takesNumbers: true, takesPoints: false, takesArray: false, invoke };
        }

        return null;
    },

    /** The method of a comparison, and of log_2 and |x|: called by the binder, not by name */
    builtIn(name, parameterCount = 2) {
        return { name, parameterCount, takesNumbers: true, takesPoints: false, takesArray: false, invoke: Functions[name] };
    },

    /** log_2, log_10: the logarithm whose base the name says after the _; null for any other name */
    logBaseOf(functionName) {
        const match = /^log_(\d{1,8})$/i.exec(functionName);
        return match == null ? null : Number(match[1]);
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
