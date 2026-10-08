// Port of Main/Avalonia/DynamicGeometry/Figures/Points/PointByCoordinates.cs: a point where
// its X and Y expressions say. Left out: the rename of expressions (the editor's).

class PointByCoordinates extends PointBase {
    constructor() {
        super();
        this.xExpression = new DrawingExpression(this, "X = ");
        this.yExpression = new DrawingExpression(this, "Y = ");
    }

    /** Always in the same order (X, then Y), which is the order the dependencies are listed in */
    get expressions() {
        return [this.xExpression, this.yExpression];
    }

    get isExpressionOwner() {
        return true;
    }

    rebindExpressions() {
        this.xExpression.rebind();
        this.yExpression.rebind();
    }

    /** The point is where its X and Y say: a drag moves it nowhere, and it keeps what is built on it from being dragged */
    allowMove() {
        return false;
    }

    recalculate() {
        // an expression that doesn't compile gives no place: the point is nowhere
        const x = this.xExpression.value;
        const y = this.yExpression.value;
        if (x == null || y == null) {
            this.exists = false;
            return;
        }

        this.coordinates = new Point(x(), y());
        this.exists = allExist(this.dependencies) && this.coordinates.exists();
    }

    onAddingToDrawing(drawing) {
        // recalculate in order to compile expressions and have accurate coordinates
        this.recalculate();
        super.onAddingToDrawing(drawing);
    }

    readXml(element) {
        super.readXml(element);
        this.xExpression.text = element.getAttribute("X");
        this.yExpression.text = element.getAttribute("Y");
    }
}

FigureTypes.register("PointByCoordinates", PointByCoordinates);
