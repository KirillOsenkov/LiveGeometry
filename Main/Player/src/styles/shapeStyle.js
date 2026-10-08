// Port of Main/Avalonia/DynamicGeometry/Styles/ShapeStyle.cs

class ShapeStyle extends LineStyle {
    static typeName = "ShapeStyle";

    constructor() {
        super();
        this.fill = new SolidColorBrush(Color.fromArgb(100, 255, 255, 200));
        this.isFilled = true;
    }

    static get properties() {
        return [["name", "string"], ["fill", "brush"], ["isFilled", "bool"], ["color", "color"], ["strokeWidth", "double"], ["dash", "enum"]];
    }

    /** [StyleFor(IShapeWithInterior)], Bezier, BezierPathInterior */
    static supportsFigure(figure) {
        return figure.isShapeWithInterior === true || figure.isBezier === true || figure.isBezierPathInterior === true;
    }

    /** The brush a shape is filled with right now: none while not filled */
    get fillBrush() {
        return this.isFilled ? this.fill : null;
    }
}
