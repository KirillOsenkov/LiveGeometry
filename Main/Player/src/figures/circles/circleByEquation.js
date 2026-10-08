// Port of Main/Avalonia/DynamicGeometry/Figures/Circles/CircleByEquation.cs: a circle given by
// expressions for its center and radius

class CircleByEquation extends CircleBase {
    constructor() {
        super();
        this.x = null;
        this.y = null;
        this.r = null;
    }

    get isShapeWithInterior() {
        return true;
    }

    get isExpressionOwner() {
        return true;
    }

    get expressions() {
        return [this.x, this.y, this.r];
    }

    rebindExpressions() {
        this.x?.rebind();
        this.y?.rebind();
        this.r?.rebind();
    }

    get center() {
        const x = this.x?.value;
        const y = this.y?.value;
        if (x == null || y == null) {
            return new Point();
        }

        return new Point(x(), y());
    }

    get radius() {
        const r = this.r?.value;
        if (r == null) {
            return 0;
        }

        const radius = r();
        return radius > 0 ? radius : GeometryMath.Epsilon;
    }

    /** A center or a radius without a value leaves no circle */
    updateExistence() {
        super.updateExistence();
        if (this.exists
            && (this.x?.value == null
                || this.y?.value == null
                || this.r?.value == null
                || !this.center.exists()
                || !(this.r.value() >= 0))) {
            this.exists = false;
        }
    }

    readXml(element) {
        this.x = new DrawingExpression(this, "Center X =", element.getAttribute("X"));
        this.y = new DrawingExpression(this, "Center Y =", element.getAttribute("Y"));
        this.r = new DrawingExpression(this, "Radius =", element.getAttribute("R"));
        super.readXml(element);
    }
}

FigureTypes.register("CircleByEquation", CircleByEquation);
