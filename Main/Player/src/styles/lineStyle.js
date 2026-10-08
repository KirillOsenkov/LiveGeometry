// Port of Main/Avalonia/DynamicGeometry/Styles/LineStyle.cs

class LineStyle extends FigureStyle {
    static typeName = "LineStyle";

    constructor() {
        super();
        this.color = Color.fromArgb(100, 0, 0, 0);
        this.strokeWidth = 1;
        this.dash = LineDash.Solid;
    }

    static get properties() {
        return [["name", "string"], ["color", "color"], ["strokeWidth", "double"], ["dash", "enum"]];
    }

    /** [StyleFor(ILinearFigure)], Slider, BezierPathPiece */
    static supportsFigure(figure) {
        return figure.isLinearFigure === true || figure.isSlider === true || figure.isBezierPathPiece === true;
    }

    /** The dash lengths in pixels for the width the line is drawn with; null for solid */
    getDashArray(strokeThickness) {
        return LineDashes.getDashArray(this.dash, strokeThickness);
    }
}
