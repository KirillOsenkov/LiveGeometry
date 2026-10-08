// Port of Main/Avalonia/DynamicGeometry/Figures/Points/Intersection/IntersectionPoint.cs

class IntersectionPoint extends PointBase {
    constructor() {
        super();

        /** One of the IntersectionAlgorithms, as the file names it */
        this.algorithm = null;
        this.algorithmName = null;
    }

    readXml(element) {
        super.readXml(element);
        const algorithm = element.getAttribute("Algorithm");
        if (algorithm == null || algorithm === "") {
            throw new Error("When reading the IntersectionPoint, the Algorithm attribute was not specified. This point will not be created.");
        }

        const method = IntersectionAlgorithms.find(algorithm);
        if (method == null) {
            throw new Error("When reading the IntersectionPoint, the Algorithm method '" + algorithm + "' wasn't found.");
        }

        this.algorithm = method;
        this.algorithmName = algorithm;
        this.recalculateAndUpdateVisual();
    }

    /** One of getAlgorithms for its two figures */
    setAlgorithm(algorithm, name) {
        this.algorithm = algorithm;
        this.algorithmName = name;
        this.recalculateAndUpdateVisual();
    }

    /**
     * Math.GetIntersectionOfCircleAndLine once swapped the two intersections when the line
     * passes through the center; a drawing from before that (IntersectionOrder="Legacy")
     * takes the other algorithm if that is the case here. Returns whether anything changed.
     */
    upgradeLegacyCircleAndLineOrder() {
        const line = this.dependencies.find(f => f.isLine === true);
        const ellipse = this.dependencies.find(f => f.isEllipse === true);
        if (line == null || ellipse == null || this.algorithm == null) {
            return false;
        }

        const name = this.algorithmName;
        const isFirst = name.endsWith("1");
        if (!isFirst && !name.endsWith("2")) {
            return false;
        }

        const center = ellipse.center;
        const projection = GeometryMath.getProjectionPoint(center, line.coordinates);
        if (!center.exists() || !projection.exists() || roundToDigits(center.distance(projection), 4) !== 0) {
            return false;
        }

        const otherName = name.substring(0, name.length - 1) + (isFirst ? "2" : "1");
        const other = IntersectionAlgorithms.find(otherName);
        if (other == null) {
            return false;
        }

        this.algorithm = other;
        this.algorithmName = otherName;
        return true;
    }

    recalculate() {
        this.exists = true;
        this.updateExistence();
        if (!this.exists) {
            return;
        }

        const figure1 = this.dependencies[0];
        const figure2 = this.dependencies[1];
        if (this.algorithm == null) {
            this.exists = false;
            return;
        }

        const p = this.algorithm(figure1, figure2);
        if (!p.exists() || figure1.hitTest(p) == null || figure2.hitTest(p) == null) {
            this.exists = false;
            return;
        }

        this.exists = true;
        this.coordinates = p;
    }

    /**
     * Every point where the two figures can cross, one algorithm each, as [name, method]:
     * one for two lines, two for a line and an ellipse or two circles, none otherwise
     */
    static getAlgorithms(figure1, figure2) {
        if (figure1.isAngleArc === true || figure2.isAngleArc === true) {
            return [];
        }

        const named = name => [name, IntersectionAlgorithms[name]];
        if (figure1.isLine === true) {
            if (figure2.isLine === true) {
                return [named("IntersectLineAndLine")];
            }

            if (figure2.isEllipse === true) {
                return [named("IntersectLineAndEllipse1"), named("IntersectLineAndEllipse2")];
            }
        } else if (figure1.isEllipse === true) {
            if (figure2.isLine === true) {
                return [named("IntersectEllipseAndLine1"), named("IntersectEllipseAndLine2")];
            }

            if (figure1.isCircle === true && figure2.isCircle === true) {
                return [named("IntersectCircleAndCircle1"), named("IntersectCircleAndCircle2")];
            }
        }

        return [];
    }
}

FigureTypes.register("IntersectionPoint", IntersectionPoint);
