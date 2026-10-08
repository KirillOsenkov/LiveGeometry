// Port of Main/Avalonia/DynamicGeometry/Figures/Points/Intersection/IntersectionAlgorithms.cs:
// the methods a file names in an IntersectionPoint's Algorithm attribute, by name.

const IntersectionAlgorithms = {
    canIntersect(figure1, figure2) {
        const lineEllipse = (figure1.isLine === true || figure1.isEllipse === true)
            && (figure2.isLine === true || figure2.isEllipse === true)
            && !(figure1.isEllipse === true && figure2.isEllipse === true);
        const lineCircle = (figure1.isLine === true || figure1.isCircle === true)
            && (figure2.isLine === true || figure2.isCircle === true);
        return lineEllipse || lineCircle;
    },

    // Line and Line

    IntersectLineAndLine(line1, line2) {
        return GeometryMath.getIntersectionOfLines(line1.coordinates, line2.coordinates);
    },

    // Circle and Line: the "Circle" names are what older files say

    IntersectCircleAndLine1(ellipse, line) {
        return IntersectionAlgorithms.IntersectEllipseAndLine(ellipse, line).p1;
    },

    IntersectCircleAndLine2(ellipse, line) {
        return IntersectionAlgorithms.IntersectEllipseAndLine(ellipse, line).p2;
    },

    IntersectLineAndCircle1(line, ellipse) {
        return IntersectionAlgorithms.IntersectEllipseAndLine1(ellipse, line);
    },

    IntersectLineAndCircle2(line, ellipse) {
        return IntersectionAlgorithms.IntersectEllipseAndLine2(ellipse, line);
    },

    // Ellipse and Line

    IntersectEllipseAndLine(ellipse, line) {
        return GeometryMath.getIntersectionOfEllipseAndLine(
            ellipse.center,
            ellipse.semiMajor,
            ellipse.semiMinor,
            ellipse.inclination,
            line.coordinates);
    },

    IntersectEllipseAndLine1(ellipse, line) {
        return IntersectionAlgorithms.IntersectEllipseAndLine(ellipse, line).p1;
    },

    IntersectEllipseAndLine2(ellipse, line) {
        return IntersectionAlgorithms.IntersectEllipseAndLine(ellipse, line).p2;
    },

    IntersectLineAndEllipse(line, ellipse) {
        return IntersectionAlgorithms.IntersectEllipseAndLine(ellipse, line);
    },

    IntersectLineAndEllipse1(line, ellipse) {
        return IntersectionAlgorithms.IntersectEllipseAndLine1(ellipse, line);
    },

    IntersectLineAndEllipse2(line, ellipse) {
        return IntersectionAlgorithms.IntersectEllipseAndLine2(ellipse, line);
    },

    // Circle and Circle

    IntersectCircleAndCircle(circle1, circle2) {
        return GeometryMath.getIntersectionOfCircles(circle1.center, circle1.radius, circle2.center, circle2.radius);
    },

    IntersectCircleAndCircle1(circle1, circle2) {
        return IntersectionAlgorithms.IntersectCircleAndCircle(circle1, circle2).p1;
    },

    IntersectCircleAndCircle2(circle1, circle2) {
        return IntersectionAlgorithms.IntersectCircleAndCircle(circle1, circle2).p2;
    },

    /** The algorithm a file names (typeof(IntersectionAlgorithms).GetMethod(name)); null when there is none */
    find(name) {
        const method = IntersectionAlgorithms[name];
        return typeof method === "function" && name !== "find" && name !== "canIntersect" ? method : null;
    }
};
