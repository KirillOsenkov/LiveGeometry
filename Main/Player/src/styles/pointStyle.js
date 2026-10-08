// Port of Main/Avalonia/DynamicGeometry/Styles/PointStyle.cs

/** What a point looks like when its style has no character (PointMarker.cs) */
const PointShape = {
    Circle: "Circle",
    Triangle: "Triangle",
    Square: "Square",
    Diamond: "Diamond",
    Pentagon: "Pentagon",
    Hexagon: "Hexagon"
};

class PointStyle extends ShapeStyle {
    static typeName = "PointStyle";

    /** A character is hard to make out at a shape's size */
    static DefaultCharacterSize = 24;

    constructor() {
        super();
        this.shape = PointShape.Circle;
        this.character = null;
        this.shapeSize = 10;
        this.characterSize = PointStyle.DefaultCharacterSize;
    }

    /** Character before Size, which depends on it: a file is read in this order */
    static get properties() {
        return [["name", "string"], ["shape", "enum"], ["character", "string"], ["size", "double"], ["fill", "brush"], ["isFilled", "bool"], ["color", "color"], ["strokeWidth", "double"], ["dash", "enum"]];
    }

    /** [StyleFor(IPoint)] */
    static supportsFigure(figure) {
        return figure.isPoint === true;
    }

    /** Drawn instead of the shape (an emoji, or any one character): null for none */
    get character() {
        return this.characterValue;
    }

    set character(value) {
        this.characterValue = value == null || value === "" ? null : value;
    }

    /** Of the shape, or the height of the character, whichever the point shows */
    get size() {
        return this.character != null ? this.characterSize : this.shapeSize;
    }

    set size(value) {
        if (this.character != null) {
            this.characterSize = value;
        } else {
            this.shapeSize = value;
        }
    }

    copyPrivateValues(source) {
        this.shapeSize = source.shapeSize;
        this.characterSize = source.characterSize;
    }
}
