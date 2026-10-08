// Port of Main/Avalonia/DynamicGeometry/Expressions/Binder.cs: resolves the names of one
// expression against a drawing, each once.

class Binder {
    constructor() {
        this.parameterName = null;
        this.drawing = null;
        this.figureAllowed = null;
        this.numbersByName = new Map();
        this.figuresByName = new Map();
        this.pointNamesValue = null;
    }

    /** The x of a function: a name that stands for the parameter, not for a figure */
    registerParameter(name) {
        this.parameterName = name;
    }

    resolveParameter(identifier) {
        return this.parameterName != null && identifier === this.parameterName ? new BoundParameter() : null;
    }

    resolveConstant(identifier) {
        if (identifier.toLowerCase() === "pi") {
            return new BoundConstant(Math.PI);
        }

        if (identifier.toLowerCase() === "e") {
            return new BoundConstant(Math.E);
        }

        return null;
    }

    /** pi, e or the x of a function; null for any other name */
    resolve(identifier) {
        return this.resolveConstant(identifier) ?? this.resolveParameter(identifier);
    }

    resolveMethod(functionName, argumentCount) {
        return ExpressionReflection.resolveMethod(functionName, argumentCount);
    }

    static takesNumbers(method, argumentCount) {
        return ExpressionReflection.takesNumbers(method, argumentCount);
    }

    /** A number of the drawing (a slider, a Number) with exactly this name, capitals and all; null when there is none */
    resolveExactNumber(name) {
        if (this.drawing == null || this.resolve(name) != null) {
            return null;
        }

        if (!this.numbersByName.has(name)) {
            const number = this.drawing.figures.list.find(f => f.isNumber === true && f.name === name) ?? null;
            this.numbersByName.set(name, number);
        }

        return this.numbersByName.get(name);
    }

    /** The names of the drawing's points, for reading two of them run together (AB) */
    get pointNames() {
        if (this.pointNamesValue == null) {
            this.pointNamesValue = this.drawing == null
                ? []
                : this.drawing.figures.list.filter(f => f instanceof PointBase).map(f => f.name);
        }

        return this.pointNamesValue;
    }

    resolveFigure(figureName) {
        if (!this.figuresByName.has(figureName)) {
            this.figuresByName.set(figureName, this.findFigure(figureName));
        }

        return this.figuresByName.get(figureName);
    }

    findFigure(figureName) {
        const figures = this.drawing.figures;
        let candidate = figures.byName(figureName);
        if (candidate == null) {
            // the indexer looks inside a composite figure (a slider) and not at it
            candidate = figures.list.find(f => f != null && f.name === figureName) ?? null;
        }

        if (candidate == null) {
            const lower = figureName.toLowerCase();
            candidate = figures.list.find(f => f != null && f.name != null && f.name !== "" && f.name.toLowerCase() === lower) ?? null;
        }

        return candidate;
    }

    isFigureAllowed(candidate) {
        if (this.figureAllowed == null) {
            return true;
        }

        return this.figureAllowed(candidate);
    }
}
