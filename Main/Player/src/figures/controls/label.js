// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/Label.cs: free text, with
// expressions, pinned to a corner of the canvas or placed in the plane.

class Label extends LabelBase {
    static BackdropPadding = 8;

    constructor() {
        super();
        this.shouldProcessText = true;
        this.pinValue = LabelPin.None;

        /** Pixels from the pinned corner of the canvas to the same corner of the label */
        this.pinOffset = new Point();
        this.backdropValue = false;
    }

    get isLengthProvider() {
        return true;
    }

    get isAngleProvider() {
        return true;
    }

    get value() {
        const exact = this.exactValue;
        if (exact != null) {
            return exact;
        }

        const text = this.processedText;
        if (Label.isNumberText(text)) {
            return Number(text.trim());
        }

        if (text === LabelBase.UndefinedText) {
            return NaN;
        }

        return 0;
    }

    static isNumberText(text) {
        return text != null && /^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?\s*$/.test(text);
    }

    /** Whether the label says a number - "[AB * 2]" - and so can stand for a length or an angle */
    get isNumber() {
        return Label.isNumberText(this.processedText);
    }

    /** For a tool that takes a figure with a length or an angle: any such figure but a label that says no number */
    static givesNumber(figure) {
        return !(figure instanceof Label) || figure.isNumber;
    }

    get angle() {
        return this.value;
    }

    get length() {
        return this.value;
    }

    // Pinning

    /**
     * A pinned label stays put on the screen while the plane zooms and pans under it: the
     * named corner of the label sits pinOffset pixels inward from the same corner of the canvas
     */
    get pin() {
        return this.pinValue;
    }

    set pin(value) {
        if (value === this.pinValue) {
            return;
        }

        if (value !== LabelPin.None && this.hasCanvas) {
            this.pinOffset = this.offsetFrom(value, this.toPhysical(this.coordinates), this.measureSize());
        }

        this.pinValue = value;

        // on the screen means on top of everything in the plane, like a control
        this.zIndex = value === LabelPin.None ? this.defaultZOrder() : ZOrder.Controls;
        if (this.hasCanvas) {
            this.updateVisual();
        }
    }

    get hasCanvas() {
        return this.drawing != null && this.drawing.canvas != null;
    }

    /** The width of the label in pixels, its text wrapping to fit; 0 lets every line run as long as it is */
    get wrapWidth() {
        return this.wrapWidthValue;
    }

    set wrapWidth(value) {
        this.wrapWidthValue = value;
        this.textLayout = null;
        if (this.drawing != null) {
            this.updateVisual();
        }
    }

    /** A plate of the paper's color behind the text, so that a caption stays readable over a grid or a figure */
    get backdrop() {
        return this.backdropValue;
    }

    set backdrop(value) {
        this.backdropValue = value;
        this.padding = value ? Label.BackdropPadding : 0;
        this.textLayout = null;
        if (this.hasCanvas) {
            this.updateVisual();
        }
    }

    /** Only a plain paper has a color to give; on a gradient there is no plate */
    get backdropBrush() {
        if (!this.backdropValue || this.drawing == null) {
            return null;
        }

        const paper = this.drawing.background;
        return paper instanceof SolidColorBrush ? paper : null;
    }

    pinnedTopLeft(size) {
        return Pinning.topLeft(this.pinValue, this.pinOffset, size, this.drawing.coordinateSystem.physicalSize);
    }

    offsetFrom(corner, topLeft, size) {
        return Pinning.offsetFrom(corner, topLeft, size, this.drawing.coordinateSystem.physicalSize);
    }

    updateVisual() {
        if (this.pinValue === LabelPin.None) {
            return;
        }

        if (!this.hasCanvas) {
            return;
        }

        const topLeft = this.pinnedTopLeft(this.measureSize());
        this.coordinates = this.toLogical(topLeft);
    }

    /** A label with live expressions depends on the figures they name, but dragging it moves the text, not those figures */
    allowMove() {
        return !this.locked;
    }

    /** Moves a pinned label on the screen, by pixels: it scrolls along with the view */
    scrollPinned(pixels) {
        if (this.pinValue === LabelPin.None || !this.hasCanvas) {
            return;
        }

        const size = this.measureSize();
        this.pinOffset = this.offsetFrom(this.pinValue, this.pinnedTopLeft(size).plus(pixels), size);
        this.updateVisual();
    }

    capturePlace() {
        return this.pinValue !== LabelPin.None ? { pinOffset: this.pinOffset } : super.capturePlace();
    }

    restorePlace(place) {
        if (place != null && place.pinOffset != null) {
            this.pinOffset = place.pinOffset;
            if (this.hasCanvas) {
                this.updateVisual();
            }
        } else {
            super.restorePlace(place);
        }
    }

    /** Dragging a pinned label changes its offset from the corner, not its place in the plane */
    moveToCore(newLocation) {
        if (this.pinValue !== LabelPin.None && this.hasCanvas) {
            this.pinOffset = this.offsetFrom(this.pinValue, this.toPhysical(newLocation), this.measureSize());
        }

        super.moveToCore(newLocation);
    }

    readXml(element) {
        super.readXml(element);
        this.text = Label.unescape(element.getAttribute("Text") ?? "");
        this.wrapWidth = Xml.readDouble(element, "WrapWidth");
        this.backdrop = Xml.readBool(element, "Backdrop", false);
        const readPin = Pinning.parse(element.getAttribute("Pin"));
        if (readPin !== LabelPin.None) {
            this.pinOffset = new Point(Xml.readDouble(element, "OffsetX"), Xml.readDouble(element, "OffsetY"));
            this.pinValue = readPin;
            this.zIndex = ZOrder.Controls;
            this.updateVisual();
        } else {
            this.moveTo(new Point(Xml.readDouble(element, "X"), Xml.readDouble(element, "Y")));
        }
    }

    /** The text as the file says it: \n is a line break, \\ a backslash */
    static unescape(text) {
        let result = "";
        for (let i = 0; i < text.length; i++) {
            if (text[i] === "\\" && i + 1 < text.length && (text[i + 1] === "n" || text[i + 1] === "\\")) {
                result += text[i + 1] === "n" ? "\n" : "\\";
                i++;
            } else {
                result += text[i];
            }
        }

        return result;
    }

    /** The processed text is compiled when the label comes into the drawing (its figures are there then) */
    onAddingToDrawing(drawing) {
        super.onAddingToDrawing(drawing);
        if (this.textChunks == null) {
            this.processText();
        }
    }
}

FigureTypes.register("Label", Label);
