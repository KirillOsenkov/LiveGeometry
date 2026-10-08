// Port of Main/Avalonia/DynamicGeometry/Figures/Points/TranslatedPoint.cs: a point at a
// distance and a direction from its source point, each tied to a figure (a vector, a length
// or an angle provider, a Number) or free (the parameter dragging changes). Left out: the
// grid's rows (Free distance...), tying and untying, the legacy format's upgrade to Numbers
// (a file from before 2026-09-25 reads, with its typed values kept as parameters).

class TranslatedPoint extends PointBase {
    constructor() {
        super();
        this.distanceQuantity = { sourceIndex: -1, parameter: 0 };
        this.directionQuantity = { sourceIndex: -1, parameter: 0 };
    }

    get source() {
        return this.dependencies.length >= 1 && this.dependencies[0].isPoint === true ? this.dependencies[0] : null;
    }

    /** A vector, a length provider or a Number; null while the distance is free */
    get distanceSource() {
        return this.sourceOf(this.distanceQuantity);
    }

    /** A vector, a line (its oriented direction), an angle provider or a Number; null while the direction is free */
    get directionSource() {
        return this.sourceOf(this.directionQuantity);
    }

    sourceOf(quantity) {
        const index = quantity.sourceIndex;
        return index >= 0 && index < this.dependencies.length ? this.dependencies[index] : null;
    }

    get isDistanceFree() {
        return this.distanceQuantity.sourceIndex < 0;
    }

    get isDirectionFree() {
        return this.directionQuantity.sourceIndex < 0;
    }

    /** Dragging changes something: this is a draggable point, styled like one */
    get hasFreedom() {
        return this.isDistanceFree || this.isDirectionFree;
    }

    /** The dependencies: the source point, then the distance source and the direction source when there are any (one entry when a vector is both) */
    setSources(source, distanceSource, directionSource) {
        const dependencies = [source];
        this.distanceQuantity.sourceIndex = -1;
        this.directionQuantity.sourceIndex = -1;
        if (distanceSource != null) {
            this.distanceQuantity.sourceIndex = dependencies.length;
            dependencies.push(distanceSource);
        }

        if (directionSource != null) {
            if (directionSource === distanceSource) {
                this.directionQuantity.sourceIndex = this.distanceQuantity.sourceIndex;
            } else {
                this.directionQuantity.sourceIndex = dependencies.length;
                dependencies.push(directionSource);
            }
        }

        this.dependencies = dependencies;
    }

    get distanceValue() {
        const source = this.distanceSource;
        if (source != null && source.isVector === true) {
            return source.length;
        }

        if (source != null && source.isLengthProvider === true) {
            return source.length;
        }

        return this.distanceQuantity.parameter;
    }

    get directionRadians() {
        const source = this.directionSource;
        if (source != null && source.isVector === true) {
            return source.direction;
        }

        // a segment, ray or line: the way it points, from its first point to its second
        if (source != null && source.isLine === true) {
            return GeometryMath.getAngle(source.coordinates.p1, source.coordinates.p2);
        }

        if (source != null && source.isAngleProvider === true) {
            return source.angle;
        }

        return this.directionQuantity.parameter;
    }

    get distance() {
        return this.distanceValue;
    }

    /** In degrees, counterclockwise from the x axis */
    get direction() {
        return GeometryMath.toDegrees(this.directionRadians);
    }

    allowMove() {
        return !this.locked && this.hasFreedom;
    }

    /** The free quantity follows the cursor: the direction as the angle from the source, the distance as the projection onto the line through the source, signed */
    moveToCore(newPosition) {
        const source = this.source;
        if (source == null) {
            return;
        }

        const origin = source.coordinates;
        if (this.isDirectionFree) {
            this.directionQuantity.parameter = GeometryMath.getAngle(origin, newPosition);
        }

        if (this.isDistanceFree) {
            const direction = this.directionRadians;
            this.distanceQuantity.parameter =
                (newPosition.x - origin.x) * Math.cos(direction)
                + (newPosition.y - origin.y) * Math.sin(direction);
        }

        this.recalculate();
    }

    capturePlace() {
        return new Point(this.distanceQuantity.parameter, this.directionQuantity.parameter);
    }

    restorePlace(place) {
        this.distanceQuantity.parameter = place.x;
        this.directionQuantity.parameter = place.y;
        this.recalculate();
        this.updateVisual();
    }

    recalculate() {
        const source = this.source;
        if (source == null || !allExist(this.dependencies)) {
            this.exists = false;
            return;
        }

        this.coordinates = GeometryMath.getTranslationPoint(source.coordinates, this.distanceValue, this.directionRadians);
        this.exists = this.coordinates.exists();
    }

    readXml(element) {
        super.readXml(element);
        const distanceSource = element.getAttribute("DistanceSource");
        const directionSource = element.getAttribute("DirectionSource");
        const newFormat = distanceSource != null
            || directionSource != null
            || element.hasAttribute("DistanceSourceIndex")
            || element.hasAttribute("DirectionSourceIndex")
            || element.hasAttribute("FreeDistance")
            || element.hasAttribute("FreeDirection");
        if (newFormat) {
            this.distanceQuantity.sourceIndex = this.indexOfSource(element, "Distance", distanceSource);
            this.directionQuantity.sourceIndex = this.indexOfSource(element, "Direction", directionSource);
            this.distanceQuantity.parameter = Xml.readDouble(element, "Distance");
            this.directionQuantity.parameter = GeometryMath.toRadians(Xml.readDouble(element, "Direction"));
        } else {
            if (this.dependencies.length > 1 && this.dependencies[1].isVector === true) {
                this.distanceQuantity.sourceIndex = 1;
                this.directionQuantity.sourceIndex = 1;
            } else {
                this.distanceQuantity.sourceIndex = this.dependencies.length > 1 && this.dependencies[1].isLengthProvider === true ? 1 : -1;
                this.directionQuantity.sourceIndex = this.dependencies.length > 2 && this.dependencies[2].isAngleProvider === true ? 2 : -1;
            }

            this.distanceQuantity.parameter = Xml.readDouble(element, "Magnitude");
            this.directionQuantity.parameter = Xml.readDouble(element, "Direction");
        }

        this.recalculate();
    }

    /** A source is named, except a part of a figure, which is said by its place among the dependencies (DistanceSourceIndex="1") */
    indexOfSource(element, quantity, name) {
        if (element.hasAttribute(quantity + "SourceIndex")) {
            const index = Math.trunc(Xml.readDouble(element, quantity + "SourceIndex"));
            return index >= 1 && index < this.dependencies.length ? index : -1;
        }

        return this.indexOfDependency(element, name);
    }

    /** The place of the named source among the figure's own Dependency elements: the name is the one the file says */
    indexOfDependency(element, name) {
        if (name == null) {
            return -1;
        }

        let index = 0;
        for (const dependency of Xml.elements(element, "Dependency")) {
            if (dependency.getAttribute("Name") === name && !dependency.hasAttribute("Part")) {
                return index < this.dependencies.length ? index : -1;
            }

            index++;
        }

        return -1;
    }
}

FigureTypes.register("TranslatedPoint", TranslatedPoint);
