// Port of Main/Avalonia/DynamicGeometry/Figures/Coordinates/CartesianGrid.cs: the grid and the
// axes, a composite the drawing always has first in its list, hidden unless the file says.

/** The grid's hidden points: the origin and the unit points the axes are built on (points by coordinates in C#) */
class GridPoint extends PointBase {
    constructor(x, y) {
        super();
        this.coordinates = new Point(x, y);
        this.visible = false;
    }

    allowMove() {
        return false;
    }

    updateExistence() {
        this.exists = true;
    }

    render(renderer) {
    }
}

class CartesianGrid extends CompositeFigure {
    constructor() {
        super();
        this.visibleValue = false;
        this.showAxesValue = true;

        // the grid's own styles, in the theme's colors (not in the drawing's style list)
        this.axisStyle = new LineStyle();
        this.axisStyle.name = "AxisStyle";
        this.axisStyle.strokeWidth = 1;
        this.gridStyle = new LineStyle();
        this.gridStyle.name = "GridStyle";
        this.gridStyle.strokeWidth = 0.5;
        this.minorGridStyle = new LineStyle();
        this.minorGridStyle.name = "MinorGridStyle";
        this.minorGridStyle.strokeWidth = 0.5;
        this.labelsStyle = new TextStyle();
        this.labelsStyle.fontSize = 12;
        this.labelsStyle.name = "LabelsStyle";
        this.refreshTheme();

        this.originPoint = new GridPoint(0, 0);
        this.xUnitPoint = new GridPoint(1, 0);
        this.yUnitPoint = new GridPoint(0, 1);
        this.originPoint.name = "Origin";
        this.xUnitPoint.name = "XUnitPoint";
        this.yUnitPoint.name = "YUnitPoint";

        this.xAxisLine = new Axis();
        this.xAxisLine.line.dependencies = [this.originPoint, this.xUnitPoint];
        this.xAxisLine.dependencies = [this.originPoint, this.xUnitPoint];
        this.yAxisLine = new Axis();
        this.yAxisLine.line.dependencies = [this.originPoint, this.yUnitPoint];
        this.yAxisLine.dependencies = [this.originPoint, this.yUnitPoint];
        this.xAxisLine.name = "XAxisLine";
        this.yAxisLine.name = "YAxisLine";
        this.axisLabels = new AxisLabelsCollection();
        this.gridLines = new RectangularGridLinesCollection();
        this.gridLines.minorStyle = this.minorGridStyle;

        this.xAxisLine.arrow.style = this.axisStyle;
        this.yAxisLine.arrow.style = this.axisStyle;
        this.gridLines.style = this.gridStyle;
        this.axisLabels.style = this.labelsStyle;

        this.children.push(
            this.originPoint,
            this.xUnitPoint,
            this.yUnitPoint,
            this.xAxisLine,
            this.yAxisLine,
            this.axisLabels,
            this.gridLines);
        this.visible = false;
    }

    /**
     * A color of the grid under the theme. On a solid paper of the drawing's own the grid
     * takes the colors of the theme whose paper is nearest to it in lightness, shifted by as
     * much as the paper differs from that theme's.
     */
    getColor(theme, color) {
        const paper = this.drawing?.getOwnBackground(theme);
        if (!(paper instanceof SolidColorBrush)) {
            return color(theme);
        }

        const lightness = c => c.r + c.g + c.b;
        const nearest = [...AppTheme.All].sort((a, b) =>
            Math.abs(lightness(a.paper) - lightness(paper.color)) - Math.abs(lightness(b.paper) - lightness(paper.color)))[0];
        const made = color(nearest);
        const shift = (value, from, to) => Math.max(0, Math.min(255, value + to - from));
        return Color.fromArgb(
            made.a,
            shift(made.r, nearest.paper.r, paper.color.r),
            shift(made.g, nearest.paper.g, paper.color.g),
            shift(made.b, nearest.paper.b, paper.color.b));
    }

    /** The paper is another, or the theme: the grid's styles read their colors again */
    refreshTheme() {
        this.axisStyle.bindToTheme("color", theme => this.getColor(theme, t => t.axis));
        this.gridStyle.bindToTheme("color", theme => this.getColor(theme, t => t.gridMajor));
        this.minorGridStyle.bindToTheme("color", theme => this.getColor(theme, t => t.gridMinor));
        this.labelsStyle.bindToTheme("color", theme => this.getColor(theme, t => t.axis));
    }

    applyStyle() {
        this.refreshTheme();
        super.applyStyle();
    }

    get visible() {
        return this.visibleValue;
    }

    set visible(value) {
        this.visibleValue = value;
        if (this.axisLabels == null) {
            return;
        }

        if (this.showAxesValue) {
            this.axisLabels.visible = value;
            this.xAxisLine.visible = value;
            this.yAxisLine.visible = value;
        }

        this.gridLines.visible = value;
        if (value && this.drawing != null) {
            this.updateVisual();
        }
    }

    get showAxes() {
        return this.showAxesValue;
    }

    set showAxes(value) {
        const shown = value && this.visibleValue;
        this.axisLabels.visible = shown;
        this.xAxisLine.visible = shown;
        this.yAxisLine.visible = shown;
        this.showAxesValue = value;
        if (shown && this.drawing != null) {
            this.updateVisual();
        }
    }

    /** Whether the axes are on screen, where a click can take them */
    get showsAxes() {
        return this.visibleValue && this.showAxesValue;
    }

    hitTestWith(point, filter) {
        return null;
    }

    hitTest(point) {
        return null;
    }

    updateVisual() {
        if (!this.visibleValue) {
            return;
        }

        super.updateVisual();
    }

    render(renderer) {
        if (!this.visibleValue) {
            return;
        }

        // the lines under the labels, the axes on top
        this.gridLines.render(renderer);
        if (this.showAxesValue) {
            this.xAxisLine.render(renderer);
            this.yAxisLine.render(renderer);
            this.axisLabels.render(renderer);
        }
    }

    toString() {
        return "Coordinate grid";
    }

    get serializable() {
        return false;
    }
}
