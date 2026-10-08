// Port of Main/Avalonia/DynamicGeometry/Figures/Points/PointOnFigure/PointOnFigure.cs

class PointOnFigure extends FreePoint {
    constructor() {
        super();

        /** Where the point is along its figure. Setting it moves nothing by itself (a locus samples through it). */
        this.parameter = 0;
        this.useHitTestingForExistence = true;
    }

    readXml(element) {
        super.readXml(element);

        // where the file says, until the drawing is worked out
        this.parameter = Xml.readDouble(element, "Parameter");
    }

    capturePlace() {
        return this.parameter;
    }

    restorePlace(place) {
        this.parameter = place;
        this.recalculateAndUpdateVisual();
    }

    get x() {
        return this.coordinates.x;
    }

    get y() {
        return this.coordinates.y;
    }

    allowMove() {
        return !this.locked;
    }

    get linearFigure() {
        return this.dependencies[0];
    }

    moveToCore(newPosition) {
        const figure = this.linearFigure;
        this.parameter = figure.getNearestParameterFromPoint(newPosition);
        newPosition = figure.getPointFromParameter(this.parameter);
        super.moveToCore(newPosition);
    }

    recalculate() {
        if (!allExist(this.dependencies)) {
            this.exists = false;
            return;
        }

        const figure1 = this.linearFigure;
        const p = figure1.getPointFromParameter(this.parameter);
        if (!p.exists()) {
            this.exists = false;
            return;
        }

        // where the parameter says, also when that place is off the figure
        this.coordinates = p;

        // a graph is hit by its samples, and those end at the edges of the window: a point
        // on it is wherever the function has a value; a locus likewise
        const hitTestFailed = this.useHitTestingForExistence
            && figure1.isFunctionGraph !== true
            && figure1.isLocus !== true
            && figure1.hitTest(p) == null;
        this.exists = !hitTestFailed;
    }

    /** Whether a click puts a new point on the figure: anything linear but the mark of an angle */
    static canBeOnFigure(figure) {
        return figure.isLinearFigure === true && figure.isAngleArc !== true;
    }
}

FigureTypes.register("PointOnFigure", PointOnFigure);
