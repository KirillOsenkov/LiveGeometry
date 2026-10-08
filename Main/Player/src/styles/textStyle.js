// Port of Main/Avalonia/DynamicGeometry/Styles/TextStyle.cs

class TextStyle extends FigureStyle {
    static typeName = "TextStyle";

    constructor() {
        super();
        this.fontSize = 10;
        this.color = Color.black;
        this.fontFamily = "Arial";
        this.bold = false;
        this.italic = false;
        this.underline = false;
    }

    static get properties() {
        return [["name", "string"], ["fontSize", "double"], ["color", "color"], ["fontFamily", "fontFamily"], ["bold", "bool"], ["italic", "bool"], ["underline", "bool"]];
    }

    /** [StyleFor(LabelBase)], ShowHideControl */
    static supportsFigure(figure) {
        return figure.isLabel === true || figure.isShowHideControl === true;
    }

    /**
     * The font as a canvas takes it. The family a drawing names (Segoe UI, Arial) is the
     * machine's it was made on: the app sets it only when it is installed, and the player
     * uses the host page's font (Fonts.family), as the app uses its own.
     */
    toCanvasFont() {
        return (this.italic ? "italic " : "") + (this.bold ? "bold " : "") + this.fontSize + "px " + Fonts.family;
    }
}

/** The font text is drawn in: the host page's, unless the player was given one */
const Fonts = {
    family: "sans-serif"
};
