// Port of Main/Avalonia/DynamicGeometry/Math.cs: the geometry. Called GeometryMath here,
// since Math is JavaScript's own. The point helpers (PointExtensions) are methods of Point
// (core/point.js). Left out: the snapping helpers of the editor (ortho, snap to grid, to
// point, to a segment's center, exact angle), IsPointInPolygonOld.

const GeometryMath = {
    Epsilon: 0.00000001,
    Precision: 0.00000001,
    PI: Math.PI,
    DOUBLEPI: 2 * Math.PI,

    get infinitePoint() {
        return Point.infinite;
    },

    get infinitePointPair() {
        return PointPair.infinite;
    },

    /** In pixels: how far from a figure a click still takes it (Settings.CursorTolerance) */
    get cursorTolerance() {
        return Settings.cursorTolerance;
    },

    set cursorTolerance(value) {
        Settings.cursorTolerance = value;
    },

    /** http://www.ecse.rpi.edu/Homepages/wrf/Research/Short_Notes/pnpoly.html */
    isPointInPolygon(polygon, start) {
        const n = polygon.length;
        const x = start.x;
        const y = start.y;
        let inside = false;
        for (let i = 0, j = n - 1; i < n; j = i++) {
            if ((polygon[i].y > y) !== (polygon[j].y > y)
                && x < (polygon[j].x - polygon[i].x) * (y - polygon[i].y) / (polygon[j].y - polygon[i].y) + polygon[i].x) {
                inside = !inside;
            }
        }

        return inside;
    },

    getIntersectionsOfPolygonAndSegment(polygon, segment, inclusive) {
        const result = [];
        const n = polygon.length;
        for (let i = 0; i < n; i++) {
            const side = new PointPair(polygon[i], polygon[(i + 1) % n]);
            const intersection = GeometryMath.getIntersectionOfSegments(side, segment, inclusive);
            if (intersection.exists()) {
                if (inclusive) {
                    result.push(intersection);
                } else if (!intersection.isEqual(polygon[(i + 1) % n])
                    && !intersection.isEqual(segment.p1)
                    && !intersection.isEqual(segment.p2)) {
                    result.push(intersection);
                }
            }
        }

        return result;
    },

    getIntersectionsOfPolygonAndLine(polygon, line, inclusive) {
        const result = [];
        const n = polygon.length;
        for (let i = 0; i < n; i++) {
            const side = new PointPair(polygon[i], polygon[(i + 1) % n]);
            const intersection = GeometryMath.getIntersectionOfSegmentAndLine(side, line, inclusive);
            if (intersection.exists()) {
                if (inclusive) {
                    if (!isWithinEpsilon(intersection.distance(side.p2))) {
                        result.push(intersection);
                    }
                } else if (!intersection.isEqual(polygon[(i + 1) % n])
                    && !intersection.isEqual(line.p1)
                    && !intersection.isEqual(line.p2)) {
                    result.push(intersection);
                }
            }
        }

        return result;
    },

    polylineLength(polyline) {
        let sum = 0;
        for (let i = 0; i < polyline.length - 1; i++) {
            sum += polyline[i].distance(polyline[i + 1]);
        }

        return sum;
    },

    getPerpendicularLine(parentLine, point) {
        return new PointPair(
            point,
            new Point(
                point.x + parentLine.p2.y - parentLine.p1.y,
                point.y + parentLine.p1.x - parentLine.p2.x));
    },

    /** Rounded for showing; see NumberFormat.round */
    round(value, fractionalDigits) {
        return NumberFormat.round(value, fractionalDigits);
    },

    scalePointBetweenTwo(p1, p2, ratio) {
        if (p1 instanceof PointPair) {
            return GeometryMath.scalePointBetweenTwo(p1.p1, p1.p2, p2);
        }

        return new Point(p1.x + (p2.x - p1.x) * ratio, p1.y + (p2.y - p1.y) * ratio);
    },

    toDegrees(radians) {
        return radians * 180 / Math.PI;
    },

    toRadians(degrees) {
        return degrees * Math.PI / 180;
    },

    /** The counterclockwise angle from vertex->first to vertex->second, in [0, 2pi) */
    oAngle(firstPoint, vertex, secondPoint) {
        const a1 = GeometryMath.getAngle(vertex, firstPoint);
        let a2 = GeometryMath.getAngle(vertex, secondPoint);
        if (a2 < a1) {
            a2 = a2 + 2 * Math.PI;
        }

        let result = a2 - a1;
        if (result > 2 * Math.PI) {
            result -= 2 * Math.PI;
        }

        // a full turn but for a rounding error is no turn
        if (2 * Math.PI - result < 1e-9) {
            result = 0;
        }

        return result;
    },

    /** The angle between two directions (GetAngle(double, double)) */
    getAngleBetween(direction1, direction2) {
        const angularSeparation = (direction2 % GeometryMath.DOUBLEPI) - (direction1 % GeometryMath.DOUBLEPI);
        return isWithinEpsilon(angularSeparation) ? 0 : angularSeparation;
    },

    /** The direction from the center to the point, in [0, 2pi) */
    getAngle(center, endPoint) {
        let result = Math.atan2(endPoint.y - center.y, endPoint.x - center.x);
        if (result < 0) {
            result += 2 * Math.PI;
        }

        return result;
    },

    getRotationPoint(p, center, angle) {
        if (angle === 0) {
            return p;
        }

        if (p.equals(center)) {
            return p;
        }

        const r = center.distance(p);
        const currentAngle = Math.atan2(p.y - center.y, p.x - center.x);
        const newAngle = currentAngle + angle;
        return new Point(center.x + r * Math.cos(newAngle), center.y + r * Math.sin(newAngle));
    },

    getSnapToGridPosition(gridSpacing, point) {
        return new Point(
            gridSpacing * Math.round(point.x / gridSpacing),
            gridSpacing * Math.round(point.y / gridSpacing));
    },

    vectorProduct(p1, p2, p3) {
        return (p2.x - p1.x) * (p3.y - p1.y) - (p3.x - p1.x) * (p2.y - p1.y);
    },

    sqr(number) {
        return number * number;
    },

    /** The square root of the absolute value (SquareRoot) */
    squareRoot(number) {
        return Math.sqrt(Math.abs(number));
    },

    /** The part of a line through the segment that is inside the borders (physical or logical, as given) */
    getLineFromSegment(segment, borders) {
        let p1 = new Point();
        let p2 = new Point();
        if (segment.p1.equals(segment.p2)) {
            return segment;
        }

        if (segment.p1.x === segment.p2.x) {
            p1 = new Point(segment.p1.x, borders.p1.y);
            p2 = new Point(segment.p1.x, borders.p2.y);
            if (segment.p1.y > segment.p2.y) {
                p1 = p1.withY(borders.p2.y);
                p2 = p2.withY(borders.p1.y);
            }
        } else if (segment.p1.y === segment.p2.y) {
            p1 = new Point(borders.p1.x, segment.p1.y);
            p2 = new Point(borders.p2.x, segment.p1.y);
            if (segment.p1.x > segment.p2.x) {
                p1 = p1.withX(borders.p2.x);
                p2 = p2.withX(borders.p1.x);
            }
        } else {
            const deltaX = segment.p2.x - segment.p1.x;
            const deltaY = segment.p2.y - segment.p1.y;
            const deltaXYRatio = deltaX / deltaY;
            const deltaYXRatio = deltaY / deltaX;

            p1 = p1.withY(deltaY > 0 ? borders.p1.y : borders.p2.y);
            p1 = p1.withX(segment.p1.x + (p1.y - segment.p1.y) * deltaXYRatio);
            if (p1.x < borders.p1.x) {
                p1 = p1.withX(borders.p1.x);
                p1 = p1.withY(segment.p1.y + (p1.x - segment.p1.x) * deltaYXRatio);
            } else if (p1.x > borders.p2.x) {
                p1 = p1.withX(borders.p2.x);
                p1 = p1.withY(segment.p1.y + (p1.x - segment.p1.x) * deltaYXRatio);
            }

            p2 = p2.withX(deltaX > 0 ? borders.p2.x : borders.p1.x);
            p2 = p2.withY(segment.p2.y + (p2.x - segment.p2.x) * deltaYXRatio);
            if (p2.y < borders.p1.y) {
                p2 = p2.withY(borders.p1.y);
                p2 = p2.withX(segment.p2.x + (p2.y - segment.p2.y) * deltaXYRatio);
            } else if (p2.y > borders.p2.y) {
                p2 = p2.withY(borders.p2.y);
                p2 = p2.withX(segment.p2.x + (p2.y - segment.p2.y) * deltaXYRatio);
            }
        }

        return new PointPair(p1, p2);
    },

    /** ProjectionInfo: the foot of the perpendicular, how far along the line it is (0 at P1, 1 at P2), and the distance */
    getProjection(point, line) {
        const projectionPoint = GeometryMath.getProjectionPoint(point, line);
        return {
            point: projectionPoint,
            ratio: GeometryMath.getProjectionRatio(line, projectionPoint),
            distanceToLine: projectionPoint.distance(point),
            isWithinSegment() {
                return this.ratio >= 0 && this.ratio <= 1;
            }
        };
    },

    getProjectionRatio(line, projection) {
        let result = 0;
        if (!isWithinEpsilon(line.p1.x - line.p2.x)) {
            result = (projection.x - line.p1.x) / (line.p2.x - line.p1.x);
        } else if (!isWithinEpsilon(line.p1.y - line.p2.y)) {
            result = (projection.y - line.p1.y) / (line.p2.y - line.p1.y);
        }

        return result;
    },

    getDistanceToLine(point, line) {
        return point.distance(GeometryMath.getProjectionPoint(point, line));
    },

    getProjectionPoint(p, line) {
        if (isWithinEpsilonTo(line.p1.y, line.p2.y)) {
            if (isWithinEpsilonTo(line.p1.x, line.p2.x)) {
                return line.p1;
            }

            return new Point(p.x, line.p1.y);
        }

        if (isWithinEpsilonTo(line.p1.x, line.p2.x)) {
            return new Point(line.p1.x, p.y);
        }

        const a = p.minus(line.p1).sumOfSquares();
        const b = p.minus(line.p2).sumOfSquares();
        const c = line.p1.minus(line.p2).sumOfSquares();
        if (c !== 0) {
            const m = (a + c - b) / (2 * c);
            return GeometryMath.scalePointBetweenTwo(line.p1, line.p2, m);
        }

        return line.p1;
    },

    /** The nearest projection onto a side of a polygonal chain (within the side) */
    getProjectionOnChain(point, polygonalChain, isClosed) {
        let nearestProjection = { point: new Point(), ratio: 0, distanceToLine: Number.MAX_VALUE, isWithinSegment() { return false; } };
        const count = polygonalChain.length;
        const stop = isClosed ? count : count - 1;
        for (let i = 0; i < stop; i++) {
            const segment = new PointPair(polygonalChain[i], polygonalChain[(i + 1) % count]);
            const projectionInfo = GeometryMath.getProjection(point, segment);
            if (projectionInfo.distanceToLine < nearestProjection.distanceToLine && projectionInfo.isWithinSegment()) {
                nearestProjection = projectionInfo;
            }
        }

        return nearestProjection;
    },

    getNearestPoint(point, pointList) {
        return pointList[GeometryMath.getIndexOfNearestPoint(point, pointList)];
    },

    getIndexOfNearestPoint(point, pointList) {
        let nearest = Infinity;
        let nearestPointIndex = 0;
        for (let i = 0; i < pointList.length; i++) {
            const currentDistance = point.distance(pointList[i]);
            if (currentDistance < nearest) {
                nearest = currentDistance;
                nearestPointIndex = i;
            }
        }

        return nearestPointIndex;
    },

    getNearestParameterFromPointOnPolyline(points, point) {
        let nearestProjectionDistance = Number.MAX_VALUE;
        let totalLength = 0;
        let parameter = 0;
        let vertexParameter = 0;
        let nearestDistance = Infinity;
        for (let i = 0; i < points.length - 1; i++) {
            const segment = new PointPair(points[i], points[i + 1]);
            const projectionInfo = GeometryMath.getProjection(point, segment);
            if (projectionInfo.distanceToLine < nearestProjectionDistance && projectionInfo.isWithinSegment()) {
                nearestProjectionDistance = projectionInfo.distanceToLine;
                parameter = totalLength + segment.length * projectionInfo.ratio;
            }

            const distance = point.distance(points[i]);
            if (distance < nearestDistance) {
                nearestDistance = distance;
                vertexParameter = totalLength;
            }

            totalLength += segment.length;
        }

        const last = points[points.length - 1];
        if (point.distance(last) < nearestDistance) {
            nearestDistance = point.distance(last);
            vertexParameter = totalLength;
        }

        if (nearestDistance < nearestProjectionDistance) {
            return vertexParameter / totalLength;
        }

        return parameter / totalLength;
    },

    isPointOnLine(line, point, epsilon) {
        return GeometryMath.getProjection(point, line).distanceToLine < epsilon;
    },

    arePointsCollinear(a, b, c) {
        if (a.equalsWithPrecision(b) || b.equalsWithPrecision(c) || c.equalsWithPrecision(a)) {
            return false;
        }

        return GeometryMath.isPointOnLine(new PointPair(a, b), c, GeometryMath.Epsilon);
    },

    isPointOnSegment(line, point, epsilon) {
        const projection = GeometryMath.getProjection(point, line);
        const slack = GeometryMath.endTolerance(line);
        return projection.distanceToLine < epsilon && projection.ratio >= -slack && projection.ratio <= 1 + slack;
    },

    /** isClosed = false for polylines and beziers, true for polygons */
    isPointOnPolygonalChain(points, point, epsilon, isClosed) {
        if (points == null || points.length === 0) {
            return false;
        }

        const projection = GeometryMath.getProjectionOnChain(point, points, isClosed);
        if (projection.distanceToLine < epsilon) {
            return true;
        }

        const nearestPoint = GeometryMath.getNearestPoint(point, points);
        return nearestPoint.distance(point) < epsilon;
    },

    isPointOnPolyline(points, point, epsilon) {
        return GeometryMath.isPointOnPolygonalChain(points, point, epsilon, false);
    },

    getAngleBisectorPoint(vertex, side1, side2) {
        const s1 = vertex.distance(side1);
        const s2 = vertex.distance(side2);
        if (s1 === 0 || s2 === 0) {
            return Point.infinite;
        }

        const a1 = vertex.angleTo(side1);
        let a2 = vertex.angleTo(side2);
        if (a2 < a1) {
            a2 += 2 * Math.PI;
        }

        const a = (a1 + a2) / 2;
        return GeometryMath.rotatePointAround(vertex, vertex.distance(side1), a);
    },

    getIntersectionOfSegments(segment1, segment2, inclusive) {
        const result = GeometryMath.getIntersectionOfLines(segment1, segment2);
        if (!result.exists()) {
            return result;
        }

        if (inclusive) {
            if (GeometryMath.isPointInSegmentBoundingRect(segment1, result) && GeometryMath.isPointInSegmentBoundingRect(segment2, result)) {
                return result;
            }
        } else if (GeometryMath.isPointInSegmentInnerBoundingRect(segment1, result) && GeometryMath.isPointInSegmentInnerBoundingRect(segment2, result)) {
            return result;
        }

        return Point.infinite;
    },

    getIntersectionOfSegmentAndLine(segment, line, inclusive) {
        const result = GeometryMath.getIntersectionOfLines(segment, line);
        if (!result.exists()) {
            return result;
        }

        if (inclusive) {
            if (GeometryMath.isPointInSegmentBoundingRect(segment, result)) {
                return result;
            }
        } else if (GeometryMath.isPointInSegmentInnerBoundingRect(segment, result)) {
            return result;
        }

        return Point.infinite;
    },

    isPointInSegmentBoundingRect(segment, point) {
        return segment.getBoundingRect().inflate(GeometryMath.Epsilon).contains(point);
    },

    isPointInSegmentInnerBoundingRect(segment, point) {
        segment = segment.getBoundingRect();
        if (isWithinEpsilonTo(segment.p1.x, segment.p2.x)) {
            return equalsWithPrecision(point.x, segment.p1.x) && point.y > segment.p1.y && point.y < segment.p2.y;
        }

        if (isWithinEpsilonTo(segment.p1.y, segment.p2.y)) {
            return equalsWithPrecision(point.y, segment.p1.y) && point.x > segment.p1.x && point.x < segment.p2.x;
        }

        return segment.containsInner(point);
    },

    getIntersectionOfLines(line1, line2) {
        const a1 = line1.p2.y - line1.p1.y;
        const b1 = line1.p1.x - line1.p2.x;
        const c1 = line1.p2.x * line1.p1.y - line1.p1.x * line1.p2.y;
        const a2 = line2.p2.y - line2.p1.y;
        const b2 = line2.p1.x - line2.p2.x;
        const c2 = line2.p2.x * line2.p1.y - line2.p1.x * line2.p2.y;
        return GeometryMath.solveLinearSystem(a1, b1, c1, a2, b2, c2);
    },

    solveLinearSystem(a1, b1, c1, a2, b2, c2) {
        const d = a1 * b2 - a2 * b1;

        // parallel by the sine of the angle between the lines, not by the determinant
        const lengths = Math.sqrt(a1 * a1 + b1 * b1) * Math.sqrt(a2 * a2 + b2 * b2);
        if (Math.abs(d) <= GeometryMath.Epsilon * lengths) {
            return Point.infinite;
        }

        const dx = b1 * c2 - b2 * c1;
        const dy = a2 * c1 - a1 * c2;
        return new Point(dx / d, dy / d);
    },

    solveSquareEquation(a, b, c) {
        const result = [];
        let d = b * b - 4 * a * c;
        if (a === 0) {
            return result;
        }

        if (d > 0) {
            d = GeometryMath.squareRoot(d);
            a *= 2;
            result.push((-b - d) / a);
            result.push((d - b) / a);
        } else if (d === 0) {
            result.push(-b / (2 * a));
        }

        return result;
    },

    /** The average of the points (Midpoint of a list; of two points) */
    midpoint(...points) {
        if (points.length === 1 && Array.isArray(points[0])) {
            points = points[0];
        }

        if (points.length === 0) {
            return new Point();
        }

        let sum = new Point();
        for (const point of points) {
            sum = sum.plus(point);
        }

        return sum.scale(1 / points.length);
    },

    /** The area of a polygon; of a circle given its center and a point on it; 0 for less */
    area(points) {
        if (points.length < 2) {
            return 0;
        }

        if (points.length === 2) {
            return GeometryMath.sqr(points[0].distance(points[1])) * Math.PI;
        }

        // a polygon whose sides cross (a bow tie) is measured as it is filled
        const levels = GeometryMath.crossingLevels(points);
        if (levels != null) {
            return GeometryMath.filledArea(points, levels);
        }

        let sum = 0;
        for (let i = 0; i < points.length - 1; i++) {
            sum += (points[i + 1].x - points[i].x) * (points[i + 1].y + points[i].y) / 2;
        }

        const lastIndex = points.length - 1;
        sum += (points[0].x - points[lastIndex].x) * (points[0].y + points[lastIndex].y) / 2;
        return Math.abs(sum);
    },

    /** The heights at which two sides of the polygon cross each other; null for a polygon whose sides don't cross */
    crossingLevels(points) {
        let levels = null;
        const count = points.length;
        for (let i = 0; i < count; i++) {
            const p = points[i];
            const r = points[(i + 1) % count].minus(p);
            for (let j = i + 2; j < count; j++) {
                if (i === 0 && j === count - 1) {
                    continue;
                }

                const q = points[j];
                const s = points[(j + 1) % count].minus(q);
                const denominator = r.x * s.y - r.y * s.x;
                if (denominator === 0) {
                    continue;
                }

                const t = ((q.x - p.x) * s.y - (q.y - p.y) * s.x) / denominator;
                const u = ((q.x - p.x) * r.y - (q.y - p.y) * r.x) / denominator;
                if (t > 0 && t < 1 && u > 0 && u < 1) {
                    levels = levels ?? [];
                    levels.push(p.y + t * r.y);
                }
            }
        }

        return levels;
    },

    /** The area an even-odd fill of the polygon covers, in slabs between the heights of its vertices and crossings */
    filledArea(points, levels) {
        for (const point of points) {
            levels.push(point.y);
        }

        levels.sort((a, b) => a - b);
        let area = 0;
        const crossings = [];
        for (let level = 1; level < levels.length; level++) {
            const height = levels[level] - levels[level - 1];
            if (!(height > 0)) {
                continue;
            }

            const middle = (levels[level] + levels[level - 1]) / 2;
            crossings.length = 0;
            for (let i = 0; i < points.length; i++) {
                const a = points[i];
                const b = points[(i + 1) % points.length];
                if ((a.y < middle) !== (b.y < middle)) {
                    crossings.push(a.x + (b.x - a.x) * (middle - a.y) / (b.y - a.y));
                }
            }

            crossings.sort((a, b) => a - b);
            for (let i = 0; i + 1 < crossings.length; i += 2) {
                area += (crossings[i + 1] - crossings[i]) * height;
            }
        }

        return area;
    },

    /** The length of a polyline through the points */
    length(points) {
        if (points.length < 2) {
            return 0;
        }

        let sum = 0;
        for (let i = 0; i < points.length - 1; i++) {
            sum += points[i].distance(points[i + 1]);
        }

        return sum;
    },

    /** The perimeter: the points joined round (Distance of a list) */
    distanceAround(points) {
        const count = points.length;
        if (count === 0) {
            return 0;
        }

        let distance = 0;
        for (let i = 0; i < count; i++) {
            distance += points[i].distance(points[(i + 1) % count]);
        }

        return distance;
    },

    getDilationPoint(p, center, factor) {
        if (factor === 0) {
            return center;
        }

        const beforeDistance = center.distance(p);
        if (beforeDistance === 0) {
            return center;
        }

        const afterDistance = beforeDistance * factor;
        const dx = p.x - center.x;
        const dy = p.y - center.y;
        return new Point(center.x + afterDistance * dx / beforeDistance, center.y + afterDistance * dy / beforeDistance);
    },

    /** How far apart two lengths near the scale may be and still count as equal when figures touch */
    tangencyTolerance(scale) {
        return 1e-9 * Math.max(1, scale);
    },

    magnitude(point) {
        return Math.max(Math.abs(point.x), Math.abs(point.y));
    },

    /**
     * How far a point is from an ellipse, along the ray from the center through the point:
     * negative inside, the radius minus the distance for a circle
     */
    radialDistanceToEllipse(center, semiMajor, semiMinor, inclination, point) {
        const distance = center.distance(point);
        if (distance === 0) {
            return -Math.min(semiMajor, semiMinor);
        }

        const angle = GeometryMath.getAngle(center, point) - inclination;
        const x = distance * Math.cos(angle);
        const y = distance * Math.sin(angle);
        const equationLeft = x * x / (semiMajor * semiMajor) + y * y / (semiMinor * semiMinor);
        return distance - distance / Math.sqrt(equationLeft);
    },

    /** How far past its ends a point may be and still be on a segment, as a fraction of its length */
    endTolerance(line) {
        return GeometryMath.tangencyTolerance(Math.max(GeometryMath.magnitude(line.p1), GeometryMath.magnitude(line.p2))) / line.length;
    },

    /** Inbound first, outbound second, by the order of the points in the line */
    getIntersectionOfCircleAndLine(center, radius, line) {
        const result = PointPair.infinite;
        const dx = line.p2.x - line.p1.x;
        const dy = line.p2.y - line.p1.y;
        const lengthSquared = dx * dx + dy * dy;
        if (lengthSquared === 0) {
            return result;
        }

        const t = ((center.x - line.p1.x) * dx + (center.y - line.p1.y) * dy) / lengthSquared;
        const p = new Point(line.p1.x + dx * t, line.p1.y + dy * t);
        const h = center.distance(p);
        const tolerance = GeometryMath.tangencyTolerance(Math.max(radius, GeometryMath.magnitude(center)));
        if (h > radius + tolerance) {
            return result;
        }

        // touching: the foot itself, exactly
        let s = 0;
        if (h < radius - tolerance) {
            s = GeometryMath.squareRoot((radius - h) * (radius + h)) / GeometryMath.squareRoot(lengthSquared);
        }

        result.p1 = new Point(p.x - dx * s, p.y - dy * s);
        result.p2 = new Point(p.x + dx * s, p.y + dy * s);
        return result;
    },

    getIntersectionOfCircleAndSegment(center, radius, segment) {
        const result = PointPair.infinite;
        const ints = GeometryMath.getIntersectionOfCircleAndLine(center, radius, segment);
        if (GeometryMath.isPointInSegmentInnerBoundingRect(segment, ints.p1)) {
            result.p1 = ints.p1;
        }

        if (GeometryMath.isPointInSegmentInnerBoundingRect(segment, ints.p2)) {
            result.p2 = ints.p2;
        }

        return result;
    },

    getIntersectionOfEllipseAndSegment(center, semiMajor, semiMinor, angle, segment) {
        const result = PointPair.infinite;
        const ints = GeometryMath.getIntersectionOfEllipseAndLine(center, semiMajor, semiMinor, angle, segment);
        if (GeometryMath.isPointInSegmentInnerBoundingRect(segment, ints.p1)) {
            result.p1 = ints.p1;
        }

        if (GeometryMath.isPointInSegmentInnerBoundingRect(segment, ints.p2)) {
            result.p2 = ints.p2;
        }

        return result;
    },

    /** Transform the line-ellipse system, treat as circle and line, then transform back */
    getIntersectionOfEllipseAndLine(center, semiMajor, semiMinor, angle, line) {
        if (semiMajor === 0 || semiMinor === 0) {
            return PointPair.infinite;
        }

        const hScale = semiMajor / semiMinor;
        const hScaleInv = semiMinor / semiMajor;
        let p1 = GeometryMath.rotatePoint(line.p1, center, -angle);
        let p2 = GeometryMath.rotatePoint(line.p2, center, -angle);
        p1 = p1.minus(center);
        p2 = p2.minus(center);
        p1 = p1.withX(p1.x * hScaleInv);
        p2 = p2.withX(p2.x * hScaleInv);
        const transformedLine = new PointPair(p1, p2);
        const ints = GeometryMath.getIntersectionOfCircleAndLine(new Point(0, 0), semiMinor, transformedLine);
        let i1 = ints.p1.withX(ints.p1.x * hScale).plus(center);
        let i2 = ints.p2.withX(ints.p2.x * hScale).plus(center);
        i1 = GeometryMath.rotatePoint(i1, center, angle);
        i2 = GeometryMath.rotatePoint(i2, center, angle);
        return new PointPair(i1, i2);
    },

    /** The same for a figure with Center, SemiMajor, SemiMinor and Inclination */
    getIntersectionOfEllipseFigureAndLine(ellipse, line) {
        return GeometryMath.getIntersectionOfEllipseAndLine(ellipse.center, ellipse.semiMajor, ellipse.semiMinor, ellipse.inclination, line);
    },

    /** The point of the ellipse at the angle t that parametrizes it (x = a cos t, y = b sin t in its own axes) */
    pointOnEllipse(center, semiMajor, semiMinor, inclination, t) {
        const cos = Math.cos(inclination);
        const sin = Math.sin(inclination);
        const x = semiMajor * Math.cos(t);
        const y = semiMinor * Math.sin(t);
        return new Point(center.x + x * cos - y * sin, center.y + x * sin + y * cos);
    },

    /** The length of the arc of the ellipse x = a cos t, y = b sin t from start over sweep: Simpson's rule over t */
    ellipseArcLength(semiMajor, semiMinor, start, sweep) {
        if (!(semiMajor > 0 && semiMinor > 0)) {
            return 0;
        }

        const intervals = 64;
        const step = sweep / intervals;
        let sum = 0;
        for (let i = 0; i <= intervals; i++) {
            const t = start + i * step;
            const speed = Math.sqrt(GeometryMath.sqr(semiMajor * Math.sin(t)) + GeometryMath.sqr(semiMinor * Math.cos(t)));
            const weight = i === 0 || i === intervals ? 1 : i % 2 === 1 ? 4 : 2;
            sum += weight * speed;
        }

        return sum * step / 3;
    },

    slope(p2, p1) {
        if (!isWithinEpsilon(p2.x - p1.x)) {
            return (p2.y - p1.y) / (p2.x - p1.x);
        }

        return NaN;
    },

    /** P1 to the right of the way from center1 to center2, P2 to the left; a touch is exact */
    getIntersectionOfCircles(center1, radius1, center2, radius2) {
        const result = PointPair.infinite;
        const dx = center2.x - center1.x;
        const dy = center2.y - center1.y;
        const distance = center1.distance(center2);
        if (distance === 0) {
            return result;
        }

        const tolerance = GeometryMath.tangencyTolerance(Math.max(
            Math.max(radius1, radius2),
            Math.max(GeometryMath.magnitude(center1), GeometryMath.magnitude(center2))));
        const outerTouch = distance - (radius1 + radius2);
        const innerTouch = distance - Math.abs(radius1 - radius2);
        if (outerTouch > tolerance || innerTouch < -tolerance) {
            return result;
        }

        let along = ((radius1 - radius2) * (radius1 + radius2) + distance * distance) / (2 * distance);
        let halfChord = 0;
        if (outerTouch < -tolerance && innerTouch > tolerance) {
            halfChord = GeometryMath.squareRoot((radius1 - along) * (radius1 + along));
            if (Number.isNaN(halfChord)) {
                halfChord = 0;
            }
        } else if (outerTouch >= -tolerance) {
            along = radius1;
        } else {
            along = radius1 >= radius2 ? radius1 : -radius1;
        }

        const ux = dx / distance;
        const uy = dy / distance;
        const foot = new Point(center1.x + ux * along, center1.y + uy * along);
        result.p1 = new Point(foot.x + uy * halfChord, foot.y - ux * halfChord);
        result.p2 = new Point(foot.x - uy * halfChord, foot.y + ux * halfChord);
        return result;
    },

    distance(p1, p2) {
        return p1.distance(p2);
    },

    getTangentPoints(outside, center, radius) {
        const distance = outside.distance(center);
        if (distance === 0 || distance < radius) {
            return PointPair.infinite;
        }

        const angle = Math.acos(radius / distance);
        const originalAngle = Math.atan2(outside.y - center.y, outside.x - center.x);
        return new PointPair(
            GeometryMath.rotatePointAround(center, radius, originalAngle + angle),
            GeometryMath.rotatePointAround(center, radius, originalAngle - angle));
    },

    getTranslationPoint(p, distance, direction) {
        if (distance === 0) {
            return p;
        }

        return new Point(p.x + distance * Math.cos(direction), p.y + distance * Math.sin(direction));
    },

    /** Rotate point about center by angle (RotatePoint(Point, Point, double)) */
    rotatePoint(point, center, angle) {
        const p = point.minus(center);
        const q = new Point(
            p.x * Math.cos(angle) - p.y * Math.sin(angle),
            p.x * Math.sin(angle) + p.y * Math.cos(angle));
        return q.plus(center);
    },

    /** The point at the angle on the circle of the radius around the center (RotatePoint(Point, double, double)) */
    rotatePointAround(center, radius, angle) {
        return new Point(center.x + radius * Math.cos(angle), center.y + radius * Math.sin(angle));
    },

    ieeeRemainder(x, y) {
        return x - y * Math.round(x / y);
    },

    isAngleBetweenAngles(a, a1, a2, clockwise) {
        if (Math.abs(GeometryMath.ieeeRemainder(a - a1, 2 * Math.PI)) < GeometryMath.Epsilon) {
            return true;
        }

        if (Math.abs(GeometryMath.ieeeRemainder(a - a2, 2 * Math.PI)) < GeometryMath.Epsilon) {
            return true;
        }

        if (isWithinEpsilon(a2 - a1)) {
            return false;
        }

        if (clockwise) {
            const temp = a1;
            a1 = a2;
            a2 = temp;
        }

        if (a2 > a1) {
            return a >= a1 && a <= a2;
        }

        if (a <= a2) {
            return true;
        }

        return a >= a1;
    },

    /** The reflection of the source through a point */
    getSymmetricPointThroughPoint(source, mirror) {
        return new Point(2 * mirror.x - source.x, 2 * mirror.y - source.y);
    },

    /** The reflection of the source across a line */
    getSymmetricPointAcrossLine(source, mirror) {
        const projection = GeometryMath.getProjectionPoint(source, mirror);
        return GeometryMath.getSymmetricPointThroughPoint(source, projection);
    },

    /** The inversion of the source in a circle */
    getSymmetricPointInCircle(source, center, radius) {
        if (radius < GeometryMath.Epsilon) {
            return Point.infinite;
        }

        const centerToDistance = source.distance(center);
        const newRadius = radius * radius / centerToDistance;
        return center.pointInDirection(source, newRadius);
    },

    getPointOnPolylineFromParameter(logicalPoints, parameter) {
        if (logicalPoints == null || logicalPoints.length === 0) {
            return Point.infinite;
        }

        let sum = 0;
        const totalLength = GeometryMath.polylineLength(logicalPoints);
        for (let i = 0; i < logicalPoints.length - 1; i++) {
            const segment = new PointPair(logicalPoints[i], logicalPoints[i + 1]);
            const oldParameter = sum / totalLength;
            sum += segment.length;
            const newParameter = sum / totalLength;
            if (newParameter > parameter) {
                const lambda = (parameter - oldParameter) / (newParameter - oldParameter);
                return GeometryMath.scalePointBetweenTwo(segment, lambda);
            }
        }

        return logicalPoints[logicalPoints.length - 1];
    },

    /** Angle from horizontal */
    oHAngle(vertex, secondPoint) {
        const firstPoint = new Point(vertex.x + 10, vertex.y);
        return -(GeometryMath.oAngle(secondPoint, vertex, firstPoint) - GeometryMath.toRadians(180));
    }
};

/** Math.BezierInfo: a cubic through four points, sampled 50 times */
class BezierInfo {
    static NumberOfPoints = 50;

    constructor(p0, p1, p2, p3) {
        this.p0 = p0;
        this.cx = 3 * (p1.x - p0.x);
        this.bx = 3 * (p2.x - p1.x) - this.cx;
        this.ax = p3.x - p0.x - this.cx - this.bx;
        this.cy = 3 * (p1.y - p0.y);
        this.by = 3 * (p2.y - p1.y) - this.cy;
        this.ay = p3.y - p0.y - this.cy - this.by;
        this.points = this.getPoints();
    }

    getPoint(t) {
        const t2 = t * t;
        const t3 = t2 * t;
        return new Point(
            this.ax * t3 + this.bx * t2 + this.cx * t + this.p0.x,
            this.ay * t3 + this.by * t2 + this.cy * t + this.p0.y);
    }

    /** The velocity along the curve at t */
    getDerivative(t) {
        return new Point(
            3 * this.ax * t * t + 2 * this.bx * t + this.cx,
            3 * this.ay * t * t + 2 * this.by * t + this.cy);
    }

    // Gauss-Legendre with five nodes on [-1, 1]
    static gaussNodes = [0, 0.5384693101056831, -0.5384693101056831, 0.9061798459386640, -0.9061798459386640];
    static gaussWeights = [0.5688888888888889, 0.4786286704993665, 0.4786286704993665, 0.2369268850561891, 0.2369268850561891];

    /** The length of the curve: the speed integrated over t, by Gauss-Legendre on eight stretches */
    get length() {
        const stretches = 8;
        let sum = 0;
        for (let i = 0; i < stretches; i++) {
            const middle = (i + 0.5) / stretches;
            const half = 0.5 / stretches;
            for (let j = 0; j < BezierInfo.gaussNodes.length; j++) {
                const velocity = this.getDerivative(middle + half * BezierInfo.gaussNodes[j]);
                sum += BezierInfo.gaussWeights[j] * half * Math.sqrt(velocity.x * velocity.x + velocity.y * velocity.y);
            }
        }

        return sum;
    }

    getPoints() {
        const result = new Array(BezierInfo.NumberOfPoints);
        const precisionMinus1 = BezierInfo.NumberOfPoints - 1;
        for (let i = 0; i < BezierInfo.NumberOfPoints; i++) {
            result[i] = this.getPoint(i / precisionMinus1);
        }

        return result;
    }

    getProjection(point) {
        return GeometryMath.getProjectionOnChain(point, this.points, false);
    }

    getNearestParameterFromPoint(point) {
        return GeometryMath.getNearestParameterFromPointOnPolyline(this.points, point);
    }
}
