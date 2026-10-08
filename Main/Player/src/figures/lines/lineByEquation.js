// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/LineByEquation.cs and ILineEquation.cs: a
// line given by expressions, y = m x + b or A x + B y + C = 0

class SlopeInterseptLineEquation {
    constructor(parent, slope, intersept) {
        this.slope = new DrawingExpression(parent, "m = ", slope);
        this.intersept = new DrawingExpression(parent, "b = ", intersept);
    }

    get lineCoordinates() {
        const slope = this.slope.value;
        const intersept = this.intersept.value;
        if (slope == null || intersept == null) {
            return new PointPair();
        }

        const m = slope();
        const b = intersept();
        if (m === 0) {
            return new PointPair(new Point(0, b), new Point(1, b));
        }

        return new PointPair(new Point(0, b), new Point(1, m + b));
    }

    get expressions() {
        return [this.slope, this.intersept];
    }
}

class GeneralFormLineEquation {
    constructor(parent, a, b, c) {
        this.a = new DrawingExpression(parent, "A = ", a);
        this.b = new DrawingExpression(parent, "B = ", b);
        this.c = new DrawingExpression(parent, "C = ", c);
    }

    get lineCoordinates() {
        const av = this.a.value;
        const bv = this.b.value;
        const cv = this.c.value;
        if (av == null || bv == null || cv == null) {
            return new PointPair();
        }

        const a = av();
        const b = bv();
        const c = cv();
        if (a === 0 && b === 0) {
            return new PointPair();
        }

        // the point of the line nearest to the origin, and a step along it
        const scale = c / (a * a + b * b);
        const x = -a * scale;
        const y = -b * scale;
        return new PointPair(new Point(x, y), new Point(x - b, y + a));
    }

    get expressions() {
        return [this.a, this.b, this.c];
    }
}

const LineEquation = {
    read(parent, element) {
        const m = element.getAttribute("m");
        const b = element.getAttribute("b");
        const A = element.getAttribute("A");
        const B = element.getAttribute("B");
        const C = element.getAttribute("C");
        if (m != null && m !== "" && b != null && b !== "") {
            return new SlopeInterseptLineEquation(parent, m, b);
        }

        if (A != null && A !== "" && B != null && B !== "" && C != null && C !== "") {
            return new GeneralFormLineEquation(parent, A, B, C);
        }

        return null;
    }
};

class LineByEquation extends LineBase {
    constructor() {
        super();
        this.equation = null;
    }

    get isLine() {
        return true;
    }

    get isExpressionOwner() {
        return true;
    }

    get expressions() {
        return this.equation != null ? this.equation.expressions : [];
    }

    rebindExpressions() {
        for (const expression of this.expressions) {
            expression.rebind();
        }
    }

    get onScreenCoordinates() {
        return GeometryMath.getLineFromSegment(this.coordinates, this.canvasLogicalBorders);
    }

    get coordinates() {
        return this.equation != null ? this.equation.lineCoordinates : new PointPair();
    }

    /** An equation without a value leaves no line */
    updateExistence() {
        super.updateExistence();
        if (!this.exists) {
            return;
        }

        const coordinates = this.equation?.lineCoordinates ?? null;
        if (coordinates == null || !coordinates.p1.exists() || !coordinates.p2.exists() || coordinates.p1.equals(coordinates.p2)) {
            this.exists = false;
        }
    }

    readXml(element) {
        this.equation = LineEquation.read(this, element);
        super.readXml(element);
    }
}

FigureTypes.register("LineByEquation", LineByEquation);
