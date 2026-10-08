// Port of Main/Avalonia/DynamicGeometry/Figures/Values/Number.cs: a number in the drawing, a
// figure without a shape that other figures depend on. Called NumberFigure here, since
// Number is JavaScript's own; the file's element is still <Number>.

class NumberFigure extends FigureBase {
    constructor() {
        super();
        this.isHitTestVisible = false;
        this.valueField = 0;
    }

    get isNumber() {
        return true;
    }

    get isLengthProvider() {
        return true;
    }

    get isAngleProvider() {
        return true;
    }

    static createAuxiliary(drawing, value) {
        const number = new NumberFigure();
        number.drawing = drawing;
        number.auxiliary = true;
        number.value = value;
        return number;
    }

    get value() {
        return this.valueField;
    }

    set value(value) {
        this.valueField = value;
        if (this.drawing != null && this.dependents.length > 0) {
            this.recalculateAllDependents();
        }
    }

    get length() {
        return this.value;
    }

    /** Angle providers speak radians; the number itself is in degrees */
    get angle() {
        return GeometryMath.toRadians(this.value);
    }

    /** n1, n2, n3 */
    generateFigureName() {
        for (let i = 1; ; i++) {
            const candidate = "n" + i;
            if (this.nameAvailable(candidate)) {
                return candidate;
            }
        }
    }

    applyStyle() {
    }

    hitTest(point) {
        return null;
    }

    readXml(element) {
        super.readXml(element);
        this.value = Xml.readDouble(element, "Value");
    }
}

FigureTypes.register("Number", NumberFigure);
