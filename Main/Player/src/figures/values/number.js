// Port of Main/Avalonia/LiveGeometry/Figures/Values/Number.cs: a number in the drawing, a
// figure other figures depend on, named a, b, c as a slider is (n1, n2 for the helpers tools
// make). Its value is typed, or worked out by a formula over the drawing (Formula), and on the
// paper it is a caption "a = 18.85" like a slider's, which a drag moves. Hidden by default: the
// helpers stay off the paper, and the file says Visible="true" for one that is shown. Called
// NumberFigure here, since Number is JavaScript's own; the file's element is still <Number>.

class NumberFigure extends CompositeFigure {
    constructor() {
        super();
        this.valueField = 0;
        this.placed = false;
        this.positionValue = new Point();
        this.showName = true;
        this.showFormula = false;
        this.expression = new DrawingExpression(this, "Formula", "");
        this.caption = new NumberCaption(this);
        this.caption.name = "number caption";
        this.addChild(this.caption);
        this.wholeHandle = new NumberWholeHandle(this);

        // the caption is a label: among the labels
        this.layer = ZOrder.Labels;
        this.visible = false;
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

    get isMovableParts() {
        return true;
    }

    get isExpressionOwner() {
        return true;
    }

    get expressions() {
        return [this.expression];
    }

    static createAuxiliary(drawing, value) {
        const number = new NumberFigure();
        number.drawing = drawing;
        number.auxiliary = true;
        number.value = value;
        return number;
    }

    // Value and formula

    /** Whether the value is typed rather than worked out by a formula */
    get isTyped() {
        return this.expression.text === "";
    }

    get value() {
        return this.valueField;
    }

    set value(value) {
        this.valueField = value;
        if (!this.isTyped) {
            this.expression.text = "";
            this.formulaChanged();
            return;
        }

        this.refreshCaption();
        if (this.drawing != null && this.dependents.length > 0) {
            this.recalculateAllDependents();
        }
    }

    /** The formula was set or cleared: the number follows it, or is typed again at the value the formula last gave */
    formulaChanged() {
        if (this.isTyped && this.dependencies.length > 0) {
            this.unregisterFromDependencies();
            this.dependencies = [];
        }

        if (this.drawing != null && this.drawing.figures.contains(this)) {
            this.recalculateAllDependents();
        }

        this.refreshCaption();
    }

    rebindExpressions() {
        this.expression.rebind();
    }

    recalculate() {
        if (!this.isTyped && this.drawing != null) {
            // a formula that doesn't compile gives no number: the figure doesn't exist, nor
            // does anything built on it; one without a value right now says "undefined"
            const worked = this.expression.value;
            if (worked == null) {
                this.exists = false;
                this.caption.exists = false;
                return;
            }

            this.valueField = worked();
        }

        this.exists = this.dependencies.every(dependency => dependency.exists);
        this.caption.exists = this.exists;
        super.recalculate();
    }

    updateExistence() {
        super.updateExistence();
        if (this.exists && !this.isTyped && this.expression.value == null) {
            this.exists = false;
        }

        this.caption.exists = this.exists;
    }

    onAddingToDrawing(drawing) {
        // compiles the formula, so that the number and its dependencies are right at once
        this.recalculate();
        super.onAddingToDrawing(drawing);
    }

    get length() {
        return this.value;
    }

    /** Angle providers speak radians; the number itself is in degrees */
    get angle() {
        return GeometryMath.toRadians(this.value);
    }

    get decimals() {
        return this.caption.decimalsToShow;
    }

    set decimals(value) {
        this.caption.decimalsToShow = value;
    }

    /** What the caption says: the name, the formula (ShowFormula) and the number, or "undefined" for a formula without a value right now */
    get captionText() {
        let text = isValidValue(this.valueField)
            ? NumberFormat.numberText(this.valueField, this.decimals)
            : LabelBase.UndefinedText;
        if (this.showFormula && !this.isTyped) {
            text = this.expression.text + " = " + text;
        }

        return this.showName ? NameDisplay.format(this.name) + " = " + text : text;
    }

    /** The caption says the number anew, where it is on the paper; a hidden one waits until it shows */
    refreshCaption() {
        if (this.drawing != null && this.visible) {
            this.caption.updateVisual();
        }
    }

    // Place

    /** Where the caption's top-left corner sits in the plane */
    get position() {
        return this.positionValue;
    }

    set position(value) {
        this.positionValue = value;
        this.placed = true;
        this.refreshCaption();
    }

    get center() {
        return this.position;
    }

    /** A number is a hidden helper until it is shown */
    get visibleByDefault() {
        return false;
    }

    get visible() {
        return super.visible;
    }

    /** Shown for the first time without a place of its own, a number goes to the middle of the window */
    set visible(value) {
        if (value && !this.placed && this.drawing != null && this.drawing.canvas != null) {
            this.position = this.drawing.coordinateSystem.viewCenter;
        }

        super.visible = value;
        if (value && this.drawing != null) {
            this.updateVisual();
        }
    }

    findMovablePart(point) {
        return this.wholeHandle;
    }

    get wholePart() {
        return this.wholeHandle;
    }

    /** A hit on the caption is a hit on the number: the caption is never handed out */
    hitTestWith(point, filter) {
        return this.visible && this.caption.hitTest(point) != null && filter(this) ? this : null;
    }

    // Name and words

    /** a, b, c from the pool sliders are named from; the helpers tools make stay n1, n2, n3 */
    generateFigureName() {
        if (this.auxiliary) {
            for (let i = 1; ; i++) {
                const candidate = "n" + i;
                if (this.nameAvailable(candidate)) {
                    return candidate;
                }
            }
        }

        const letters = "abcdfghkmnpqrstuvwz";
        for (let i = 0; ; i++) {
            for (const letter of letters) {
                const candidate = i === 0 ? letter : letter + i;
                if (this.nameAvailable(candidate)) {
                    return candidate;
                }
            }
        }
    }

    get name() {
        return super.name;
    }

    set name(value) {
        super.name = value;
        this.refreshCaption();
    }

    toString() {
        return this.name;
    }

    /** The caption's text style stands for the number's */
    get style() {
        return this.caption.style;
    }

    set style(value) {
        this.caption.style = value;
    }

    readXml(element) {
        super.readXml(element);
        if (element.hasAttribute("Decimals")) {
            this.decimals = Math.trunc(Xml.readDouble(element, "Decimals"));
        }

        this.showName = Xml.readBool(element, "ShowName", true);
        this.showFormula = Xml.readBool(element, "ShowFormula", false);
        if (element.hasAttribute("X")) {
            this.position = new Point(Xml.readDouble(element, "X"), Xml.readDouble(element, "Y"));
        }

        const formula = element.getAttribute("Formula");
        if (formula != null && formula !== "") {
            this.expression.text = formula;
        } else {
            this.value = Xml.readDouble(element, "Value");
        }
    }
}

/** Dragging the caption moves the number by as much as the cursor, the text keeping its grip under it */
class NumberWholeHandle {
    constructor(number) {
        this.number = number;
    }

    get isMovable() {
        return true;
    }

    get coordinates() {
        return this.number.position;
    }

    allowMove() {
        return !this.number.locked;
    }

    moveTo(position) {
        this.number.position = position;
    }

    capturePlace() {
        return this.number.position;
    }

    restorePlace(place) {
        this.number.position = place;
    }
}

/** "a = 18.85" with its top-left corner at the number's place. Not draggable on its own: a drag moves the number. */
class NumberCaption extends LabelWithOffset {
    constructor(number) {
        super();
        this.number = number;
    }

    get anchor() {
        return this.number.position;
    }

    get hasAnchor() {
        return true;
    }

    allowMove() {
        return false;
    }

    updateVisual() {
        if (this.drawing == null) {
            return;
        }

        this.setText(this.number.captionText);
        super.updateVisual();
    }
}

FigureTypes.register("Number", NumberFigure);
