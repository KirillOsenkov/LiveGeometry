// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/FunctionGraph/FunctionGraph.cs: the
// graph of y = f(x) over the window, with gaps where the function has no value or jumps.

class FunctionGraph extends Curve {
    /** Units up for one across: a step of the graph steeper than this is looked into */
    static SteepSlope = 8;

    /** In pixels: a step of the graph that climbs less than this is no jump to look for */
    static JumpPixels = 24;

    constructor() {
        super();
        this.mFunction = null;
        this.functionTextValue = null;
    }

    get isFunctionGraph() {
        return true;
    }

    /** A sample every couple of pixels */
    get stepCount() {
        if (this.drawing == null || this.drawing.coordinateSystem == null) {
            return 0;
        }

        return Math.trunc(this.drawing.coordinateSystem.physicalSize.x / 2);
    }

    get function() {
        return this.mFunction;
    }

    readXml(element) {
        super.readXml(element);
        this.functionTextValue = element.getAttribute("Function");
    }

    get functionText() {
        return this.functionTextValue;
    }

    set functionText(value) {
        this.functionTextValue = value;
        this.compile();
    }

    rebindExpressions() {
        if (this.functionTextValue != null && this.functionTextValue !== "") {
            this.compile();
        }
    }

    compile() {
        const result = Compiler.instance.compileFunction(this.drawing, this.functionText, figure => !figure.dependsOn(this));
        if (result.isSuccess) {
            this.setFunction(result);
        }

        return result;
    }

    recalculate() {
        if (this.function == null) {
            this.compile();
        }

        super.recalculate();
    }

    setFunction(result) {
        this.mFunction = result.function;
        this.unregisterFromDependencies();
        this.dependencies = [...result.dependencies];

        // see DrawingExpression.recalculate
        if (this.drawing.figures.contains(this)) {
            this.registerWithDependencies();
            this.recalculateAllDependents();
        }
    }

    /** No value where there is no function, and where it throws */
    callFunction(parameter) {
        if (this.mFunction == null) {
            return NaN;
        }

        try {
            return this.mFunction(parameter);
        } catch (error) {
            return NaN;
        }
    }

    getPoints(result) {
        const stepCount = this.stepCount;
        if (stepCount === 0 || this.function == null) {
            return;
        }

        const coordinates = this.drawing.coordinateSystem;
        const minX = coordinates.minimalVisibleX;
        const maxX = coordinates.maximalVisibleX;

        // far beyond the window one value is as good as another
        const height = coordinates.maximalVisibleY - coordinates.minimalVisibleY;
        const lowest = coordinates.minimalVisibleY - 10 * height;
        const highest = coordinates.maximalVisibleY + 10 * height;
        const clamped = (x, y) => new Point(x, Math.max(lowest, Math.min(highest, y)));

        let previousX = 0;
        let previousY = NaN;
        for (let i = 0; i <= stepCount; i++) {
            const x = i === stepCount ? maxX : minX + (maxX - minX) * i / stepCount;
            const y = this.callFunction(x);
            if (!isValidValue(y)) {
                // the graph goes on to where the function stops having a value
                if (i > 0 && isValidValue(previousY)) {
                    const edge = this.findEdge(previousX, x);
                    result.push(clamped(edge, this.callFunction(edge)));
                }

                FunctionGraph.addGap(result);
            } else {
                if (i > 0 && !isValidValue(previousY)) {
                    const edge = this.findEdge(x, previousX);
                    result.push(clamped(edge, this.callFunction(edge)));
                } else if (isValidValue(previousY)
                    && FunctionGraph.mayBeJump(previousY, y, coordinates)
                    && this.isJump(previousX, previousY, x, y)) {
                    FunctionGraph.addGap(result);
                }

                result.push(clamped(x, y));
            }

            previousX = x;
            previousY = y;
        }
    }

    /** Between a place where the function has a value and one where it has none: the last place where it has */
    findEdge(inside, outside) {
        for (let i = 0; i < 20; i++) {
            const middle = (inside + outside) / 2;
            if (isValidValue(this.callFunction(middle))) {
                inside = middle;
            } else {
                outside = middle;
            }
        }

        return inside;
    }

    static addGap(points) {
        if (points.length > 0 && points[points.length - 1].exists()) {
            points.push(Curve.Gap);
        }
    }

    /** Whether the line between two samples next to each other is worth looking into: long on the screen, and not wholly above or below the window */
    static mayBeJump(y1, y2, coordinates) {
        if (Math.abs(y2 - y1) * coordinates.unitLength < FunctionGraph.JumpPixels) {
            return false;
        }

        const bothAbove = y1 > coordinates.maximalVisibleY && y2 > coordinates.maximalVisibleY;
        const bothBelow = y1 < coordinates.minimalVisibleY && y2 < coordinates.minimalVisibleY;
        return !bothAbove && !bothBelow;
    }

    /** Whether the function jumps between two samples rather than climbs: halving the step towards the larger change, a climb gets smaller and a jump stays */
    isJump(x1, y1, x2, y2) {
        const change = Math.abs(y2 - y1);
        if (change <= FunctionGraph.SteepSlope * (x2 - x1)) {
            return false;
        }

        for (let i = 0; i < 12; i++) {
            const middle = (x1 + x2) / 2;
            const y = this.callFunction(middle);
            if (!isValidValue(y)) {
                return true;
            }

            if (Math.abs(y - y1) > Math.abs(y2 - y)) {
                x2 = middle;
                y2 = y;
            } else {
                x1 = middle;
                y1 = y;
            }
        }

        return Math.abs(y2 - y1) > change / 8;
    }

    getNearestParameterFromPoint(point) {
        return point.x;
    }

    getPointFromParameter(parameter) {
        return new Point(parameter, this.callFunction(parameter));
    }

    getParameterDomain() {
        const coordinates = this.drawing.coordinateSystem;
        return [coordinates.minimalVisibleX, coordinates.maximalVisibleX];
    }
}

FigureTypes.register("FunctionGraph", FunctionGraph);
