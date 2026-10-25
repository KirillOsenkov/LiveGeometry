// Port of Main/Avalonia/LiveGeometry/Styles/ShapeStyle.cs

class ShapeStyle extends LineStyle {
    static typeName = "ShapeStyle";

    constructor() {
        super();
        this.fill = new SolidColorBrush(Color.fromArgb(100, 255, 255, 200));
        this.isFilled = true;
        this.texture = FillTexture.None;
    }

    static get properties() {
        return [["name", "string"], ["fill", "brush"], ["isFilled", "bool"], ["texture", "enum"], ["color", "color"], ["strokeWidth", "double"], ["dash", "enum"]];
    }

    /** [StyleFor(IShapeWithInterior)], Bezier, BezierPathInterior, AngleArc (which fills the angle under its mark) */
    static supportsFigure(figure) {
        return figure.isShapeWithInterior === true || figure.isBezier === true || figure.isBezierPathInterior === true || figure.isAngleArc === true;
    }

    /** The brush a shape is filled with right now: none while not filled, the fill's color laid down with the texture when the style has one */
    get fillBrush() {
        if (!this.isFilled) {
            return null;
        }

        if ((this.texture === FillTexture.Crayon || this.texture === FillTexture.FeltTip) && this.fill != null) {
            return new TexturedBrush(TexturedBrush.colorOf(this.fill), this.texture);
        }

        return this.fill;
    }
}
