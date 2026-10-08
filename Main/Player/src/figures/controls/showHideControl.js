// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/ShowHideControl.cs: a check box on
// the paper that shows and hides figures: its dependencies, which it is not built on but
// holds. The check box is drawn by the renderer (a box and a caption; Fluent's in the app).
// Left out: Edit figures, the undo of a click.

class ShowHideControl extends ControlBase {
    /** The box's side in pixels, the gap to the caption, and the control's least height (Fluent's CheckBox) */
    static BoxSize = 20;
    static Gap = 8;
    static MinHeight = 32;

    constructor() {
        super();
        this.isChecked = true;
        this.textValue = "";
        this.pin = LabelPin.None;

        /** Pixels from the pinned corner of the canvas to the same corner of the box */
        this.pinOffset = new Point();
        this.textLayout = null;
    }

    get isShowHideControl() {
        return true;
    }

    /** The caption beside the box */
    get text() {
        return this.textValue;
    }

    set text(value) {
        this.textValue = value ?? "";
        this.textLayout = null;
        if (this.drawing != null) {
            this.updateVisual();
        }
    }

    /** A box is dragged by itself: what it holds is not what it is built on */
    allowMove() {
        return !this.locked;
    }

    get font() {
        const style = this.resolvedStyle;
        return style != null && style.toCanvasFont != null ? style.toCanvasFont() : Fonts.family;
    }

    get textColor() {
        return this.resolvedStyle?.color ?? Color.black;
    }

    getTextLayout() {
        const canvas = this.canvas;
        if (canvas == null) {
            return null;
        }

        const font = this.font;
        if (this.textLayout == null || this.textLayout.text !== this.textValue || this.textLayout.font !== font) {
            this.textLayout = canvas.measureText(this.textValue, font, 0);
        }

        return this.textLayout;
    }

    /** The size of the box in pixels, measured now */
    measureSize() {
        const layout = this.getTextLayout();
        const textWidth = layout != null ? layout.width : 0;
        const textHeight = layout != null ? layout.height : 0;
        return new Size(
            ShowHideControl.BoxSize + (textWidth > 0 ? ShowHideControl.Gap + textWidth : 0),
            Math.max(ShowHideControl.MinHeight, textHeight));
    }

    readXml(element) {
        super.readXml(element);

        // only the box: each figure says in the file whether it is hidden
        this.setBox(Xml.readBool(element, "Show", true));
        this.textValue = element.getAttribute("Text") ?? "";
        const readPin = Pinning.parse(element.getAttribute("Pin"));
        if (readPin !== LabelPin.None) {
            this.pinOffset = new Point(Xml.readDouble(element, "OffsetX"), Xml.readDouble(element, "OffsetY"));
            this.pin = readPin;
            this.updateVisual();
        } else {
            this.moveTo(new Point(Xml.readDouble(element, "X"), Xml.readDouble(element, "Y")));
        }
    }

    get hasCanvas() {
        return this.drawing != null && this.drawing.canvas != null;
    }

    pinnedTopLeft() {
        return Pinning.topLeft(this.pin, this.pinOffset, this.measureSize(), this.drawing.coordinateSystem.physicalSize);
    }

    updateVisual() {
        if (this.pin === LabelPin.None) {
            return;
        }

        if (!this.hasCanvas) {
            return;
        }

        this.coordinates = this.toLogical(this.pinnedTopLeft());
    }

    /** Dragging a pinned box changes its offset from the corner, not its place in the plane */
    moveToCore(newLocation) {
        if (this.pin !== LabelPin.None && this.hasCanvas) {
            this.pinOffset = Pinning.offsetFrom(this.pin, this.toPhysical(newLocation), this.measureSize(), this.drawing.coordinateSystem.physicalSize);
        }

        super.moveToCore(newLocation);
    }

    capturePlace() {
        return this.pin !== LabelPin.None ? { pinOffset: this.pinOffset } : super.capturePlace();
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

    /** A click on the box that reached the Drag tool: ticks or unticks it */
    click() {
        const show = !this.isChecked;
        this.setBox(show);
        this.show(show);
    }

    /** Ticks or unticks the box, leaving its figures as they are */
    setBox(isChecked) {
        this.isChecked = isChecked;
    }

    updateFigureVisibility() {
        this.show(this.isChecked);
    }

    show(show) {
        for (const figure of this.dependencies) {
            figure.visible = show;
            figure.updateVisual();
        }
    }

    /** The box stays whether or not its figures exist right now: they are what it shows and hides, not what it is built on */
    updateExistence() {
        this.exists = true;
    }

    render(renderer) {
        if (!this.isShown) {
            return;
        }

        const topLeft = this.toPhysical(this.coordinates);
        if (!topLeft.exists()) {
            return;
        }

        renderer.drawCheckBox(topLeft, this.measureSize(), this.isChecked, this.getTextLayout(), this.font, this.textColor, ShowHideControl.BoxSize, ShowHideControl.Gap);
    }
}

FigureTypes.register("ShowHideControl", ShowHideControl);
