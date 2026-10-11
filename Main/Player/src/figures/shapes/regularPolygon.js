// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/RegularPolygon.cs: a regular polygon on a
// center and a vertex. Left out: the side's length as a setting, Fix length.

class RegularPolygon extends DependentPolygonBase {
    static DefaultNumberOfSides = 5;

    constructor() {
        super();
        this.numberOfSidesValue = RegularPolygon.DefaultNumberOfSides;
    }

    get isPolygon() {
        return true;
    }

    get centerPoint() {
        return this.dependencies[0];
    }

    get vertexPoint() {
        return this.dependencies[1];
    }

    get radiusToSide() {
        return 2 * Math.sin(Math.PI / this.numberOfSides);
    }

    /** The side, which is the vertex's distance from the center scaled by the number of sides */
    get length() {
        return this.center.distance(this.vertex) * this.radiusToSide;
    }

    get numberOfSides() {
        return this.numberOfSidesValue;
    }

    set numberOfSides(value) {
        if (value < 3 || value > 500) {
            return;
        }

        const minimum = this.minimumNumberOfSides();
        if (value < minimum) {
            value = minimum;
        }

        this.numberOfSidesValue = value;
        this.recreate(value);
        this.recalculateAllDependents();
    }

    /** The fewest sides that keep every vertex and side something outside the polygon is built on */
    minimumNumberOfSides() {
        let minimum = 3;
        for (let i = 0; i < this.vertices.length; i++) {
            if (i + 2 > minimum && this.hasOutsideDependents(this.vertices[i])) {
                minimum = i + 2;
            }
        }

        for (let i = 0; i < this.sides.length; i++) {
            if (i + 1 > minimum && this.hasOutsideDependents(this.sides[i])) {
                minimum = i + 1;
            }
        }

        return minimum;
    }

    readXml(element) {
        super.readXml(element);
        const count = Math.trunc(Xml.readDouble(element, "Sides"));
        if (count >= 3 && count <= 500) {
            this.numberOfSidesValue = count;
        }

        // the parts now, not at the first recalculation: a figure of the file built on a vertex looks for it as soon as it is read
        if (this.drawing != null) {
            this.recreate(this.numberOfSidesValue, false);
            this.readPartStyles(element);
        }
    }

    get center() {
        return this.point(0);
    }

    get vertex() {
        return this.point(1);
    }

    recalculate() {
        if (this.sides.length !== this.numberOfSides) {
            this.recreate(this.numberOfSides);
            return;
        }

        const center = this.center;
        const vertex = this.vertex;
        const initialAngle = GeometryMath.getAngle(center, vertex);
        const radius = center.distance(vertex);
        const increment = GeometryMath.DOUBLEPI / this.numberOfSides;
        for (let i = 0; i < this.numberOfSides - 1; i++) {
            const angle = initialAngle + (i + 1) * increment;
            this.vertices[i].moveTo(new Point(center.x + radius * Math.cos(angle), center.y + radius * Math.sin(angle)));
        }

        this.updateVisual();
    }

    collectPolygonDependencies(callback) {
        callback(this.dependencies[1]);
        super.collectPolygonDependencies(callback);
    }

    adjustVerticesList(sideCount) {
        if (this.vertices.length < sideCount - 1) {
            const requiredNumber = sideCount - this.vertices.length - 1;
            for (let i = 0; i < requiredNumber; i++) {
                this.addVertex();
            }
        } else if (this.vertices.length >= sideCount) {
            const requiredNumber = this.vertices.length - sideCount;
            for (let i = 0; i <= requiredNumber; i++) {
                this.removeVertex();
            }
        }
    }

    addSide(sideCount) {
        const isNew = this.retiredSides.length === 0;
        const side = isNew ? new PolygonSide(this) : this.retiredSides.pop();
        const common = isNew ? DependentPolygonBase.commonStyle(this.sides) : null;
        side.drawing = this.drawing;
        side.visible = this.visible;
        const index = this.sides.length;
        if (index > 2) {
            const firstSide = this.sides[index - 1];
            firstSide.unregisterFromDependencies();
            firstSide.dependencies = [firstSide.dependencies[0], this.vertices[index - 1]];
            this.registerPart(firstSide);
        }

        if (index === 0) {
            side.dependencies = [this.dependencies[1], this.vertices[0]];
        } else if (index === sideCount - 1) {
            side.dependencies = [this.vertices[sideCount - 2], this.dependencies[1]];
        } else {
            side.dependencies = [this.vertices[index - 1], this.vertices[index]];
        }

        side.z = this.z;
        this.sides.push(side);
        this.children.push(side);
        if (this.isOnCanvas) {
            side.onAddingToCanvas(this.drawing.canvas);
        }

        if (common != null) {
            side.style = common;
        }

        this.registerPart(side);
    }

    removeSide() {
        const index = this.sides.length - 1;
        if (index > 2) {
            const firstSide = this.sides[index - 1];
            firstSide.unregisterFromDependencies();
            firstSide.dependencies = [firstSide.dependencies[0], this.dependencies[1]];
            this.registerPart(firstSide);
        }

        const side = this.sides.pop();
        this.removePart(side);
        this.retiredSides.push(side);
    }

    toString() {
        return this.name;
    }
}

FigureTypes.register("RegularPolygon", RegularPolygon);
