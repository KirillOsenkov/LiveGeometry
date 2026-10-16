// Port of Main/Avalonia/LiveGeometry/Figures/Points/Intersection/CurveIntersections.cs:
// where a function graph or a Bezier curve or path crosses a line, a circle, an ellipse or a
// graph. The point keeps where on the curve it is (its parameter) and takes the crossing
// nearest that each time it is worked out, a root of the other figure's equation along the
// curve, looked for outwards on both sides and narrowed down by halving.

const CurveIntersections = (() => {
    const firstStep = 1e-3;
    const growth = 1.1;
    const maxSteps = 250;
    const rootTolerance = 1e-7;
    const nowhere = () => new Point(NaN, NaN);

    function holder(figure) {
        return figure != null && figure.isBezierPathPiece === true ? figure.owner : figure;
    }

    function isBezier(figure) {
        return figure.isBezier === true || figure.isBezierPath === true;
    }

    function hasEquation(figure) {
        return (figure.isLine === true || figure.isEllipse === true || figure.isFunctionGraph === true) && figure.isAngleArc !== true;
    }

    function findCurve(first, second) {
        first = holder(first);
        second = holder(second);
        if (first == null || second == null || first === second) {
            return null;
        }

        if (isBezier(first) && hasEquation(second)) {
            return first;
        }

        if (isBezier(second) && hasEquation(first)) {
            return second;
        }

        if (first.isFunctionGraph === true && hasEquation(second)) {
            return first;
        }

        if (second.isFunctionGraph === true && hasEquation(first)) {
            return second;
        }

        return null;
    }

    function other(curve, first, second) {
        return holder(first) === curve ? holder(second) : holder(first);
    }

    function pointAt(curve, parameter) {
        if (curve.isFunctionGraph === true) {
            return curve.getPointFromParameter(parameter);
        }

        if (curve.isBezier === true) {
            const info = curve.curveInfo;
            return info != null && info.points != null && parameter >= 0 && parameter <= 1 ? info.getPoint(parameter) : nowhere();
        }

        if (curve.isBezierPath === true) {
            return curve.getPointFromParameter(parameter);
        }

        return nowhere();
    }

    function equation(figure, point) {
        if (figure.isLine === true) {
            const ends = figure.coordinates;
            const dx = ends.p2.x - ends.p1.x;
            const dy = ends.p2.y - ends.p1.y;
            const length = Math.sqrt(dx * dx + dy * dy);
            if (!(length > 0)) {
                return NaN;
            }

            return ((point.x - ends.p1.x) * dy - (point.y - ends.p1.y) * dx) / length;
        }

        if (figure.isCircle === true) {
            return figure.radius > 0 ? figure.center.distance(point) - figure.radius : NaN;
        }

        if (figure.isEllipse === true) {
            const a = figure.semiMajor;
            const b = figure.semiMinor;
            if (!(a > 0 && b > 0)) {
                return NaN;
            }

            const center = figure.center;
            const cos = Math.cos(figure.inclination);
            const sin = Math.sin(figure.inclination);
            const dx = point.x - center.x;
            const dy = point.y - center.y;
            const u = (dx * cos + dy * sin) / a;
            const v = (-dx * sin + dy * cos) / b;
            return (Math.sqrt(u * u + v * v) - 1) * Math.min(a, b);
        }

        if (figure.isFunctionGraph === true) {
            return point.y - figure.getPointFromParameter(point.x).y;
        }

        return NaN;
    }

    class RootSearch {
        constructor(fn, curvePoint) {
            this.fn = fn;
            this.curvePoint = curvePoint;
        }

        value(t) {
            const value = this.fn(t);
            return isValidValue(value) ? value : NaN;
        }

        isRoot(t) {
            const value = this.value(t);
            if (Number.isNaN(value)) {
                return false;
            }

            const point = this.curvePoint(t);
            const scale = 1 + Math.abs(point.x) + Math.abs(point.y);
            return Math.abs(value) <= rootTolerance * scale;
        }

        findNearest(start, step, low, high) {
            const startValue = this.value(start);
            if (startValue === 0) {
                return start;
            }

            const right = { direction: 1, previousT: NaN, previousValue: NaN, lastT: start, lastValue: startValue, done: false };
            const left = { direction: -1, previousT: NaN, previousValue: NaN, lastT: start, lastValue: startValue, done: false };
            let distance = 0;
            for (let i = 0; i < maxSteps && !(right.done && left.done); i++) {
                distance += step;
                step *= growth;
                let best = NaN;
                const consider = candidate => {
                    if (!Number.isNaN(candidate) && (Number.isNaN(best) || Math.abs(candidate - start) < Math.abs(best - start))) {
                        best = candidate;
                    }
                };

                consider(this.advance(right, start + distance, low, high));
                consider(this.advance(left, start - distance, low, high));
                if (i === 0) {
                    // the start itself the lowest of the three: a touch there, or the crossing found last time
                    consider(this.touch(left.lastT, left.lastValue, start, startValue, right.lastT, right.lastValue));
                }

                if (!Number.isNaN(best)) {
                    return best;
                }
            }

            return NaN;
        }

        advance(side, t, low, high) {
            if (side.done) {
                return NaN;
            }

            if (t <= low || t >= high) {
                t = t <= low ? low : high;
                side.done = true;
                if (t === side.lastT) {
                    return NaN;
                }
            }

            const value = this.value(t);
            let result = NaN;
            const lastValid = !Number.isNaN(side.lastValue);
            const valid = !Number.isNaN(value);
            if (valid && value === 0) {
                result = t;
            } else if (lastValid && valid) {
                if ((side.lastValue < 0) !== (value < 0)) {
                    result = this.halve(side.lastT, side.lastValue, t, value);
                } else {
                    result = this.touch(side.previousT, side.previousValue, side.lastT, side.lastValue, t, value);
                }
            } else if (lastValid !== valid) {
                // the function stops having a value in between: sqrt(x) meets the x-axis at its very end
                const edge = valid ? this.findEdge(t, side.lastT) : this.findEdge(side.lastT, t);
                if (this.isRoot(edge)) {
                    result = edge;
                }
            }

            side.previousT = side.lastT;
            side.previousValue = side.lastValue;
            side.lastT = t;
            side.lastValue = value;
            return result;
        }

        halve(a, valueA, b, valueB) {
            for (let i = 0; i < 200; i++) {
                const middle = (a + b) / 2;
                if (middle === a || middle === b) {
                    break;
                }

                const value = this.value(middle);
                if (Number.isNaN(value)) {
                    return NaN;
                }

                if (value === 0) {
                    return middle;
                }

                if ((value < 0) === (valueA < 0)) {
                    a = middle;
                    valueA = value;
                } else {
                    b = middle;
                    valueB = value;
                }
            }

            const root = Math.abs(valueA) <= Math.abs(valueB) ? a : b;
            return this.isRoot(root) ? root : NaN;
        }

        touch(a, valueA, middle, valueMiddle, b, valueB) {
            if (Number.isNaN(valueA) || Number.isNaN(valueMiddle) || Number.isNaN(valueB)
                || (valueA < 0) !== (valueMiddle < 0) || (valueB < 0) !== (valueMiddle < 0)
                || !(Math.abs(valueMiddle) < Math.abs(valueA))
                || !(Math.abs(valueMiddle) <= Math.abs(valueB))) {
                return NaN;
            }

            if (a > b) {
                [a, b] = [b, a];
            }

            // golden section on the size of the value
            const ratio = 0.6180339887498949;
            let x1 = b - ratio * (b - a);
            let x2 = a + ratio * (b - a);
            let f1 = this.size(x1);
            let f2 = this.size(x2);
            for (let i = 0; i < 100 && x1 < x2; i++) {
                if (f1 <= f2) {
                    b = x2;
                    x2 = x1;
                    f2 = f1;
                    x1 = b - ratio * (b - a);
                    f1 = this.size(x1);
                } else {
                    a = x1;
                    x1 = x2;
                    f1 = f2;
                    x2 = a + ratio * (b - a);
                    f2 = this.size(x2);
                }
            }

            const lowest = f1 <= f2 ? x1 : x2;
            return this.isRoot(lowest) ? lowest : NaN;
        }

        size(t) {
            const value = this.value(t);
            return Number.isNaN(value) ? Infinity : Math.abs(value);
        }

        findEdge(inside, outside) {
            for (let i = 0; i < 100; i++) {
                const middle = (inside + outside) / 2;
                if (middle === inside || middle === outside) {
                    break;
                }

                if (Number.isNaN(this.value(middle))) {
                    outside = middle;
                } else {
                    inside = middle;
                }
            }

            return inside;
        }
    }

    return {
        holder,
        findCurve,

        canIntersect(first, second) {
            return findCurve(first, second) != null;
        },

        /** The crossing nearest the parameter on the curve, as { point, parameter }; point nowhere when there is none */
        find(first, second, parameter) {
            const curve = findCurve(first, second);
            if (curve == null) {
                return { point: nowhere(), parameter };
            }

            const otherFigure = other(curve, first, second);
            let low = -Infinity;
            let high = Infinity;
            if (curve.isFunctionGraph !== true) {
                [low, high] = curve.getParameterDomain();
            }

            let start = isValidValue(parameter) ? parameter : 0;
            start = Math.max(low, Math.min(high, start));
            const step = curve.isFunctionGraph === true ? firstStep * Math.max(1, Math.abs(start)) : firstStep;
            const search = new RootSearch(t => equation(otherFigure, pointAt(curve, t)), t => pointAt(curve, t));
            const root = search.findNearest(start, step, low, high);
            if (!isValidValue(root)) {
                return { point: nowhere(), parameter };
            }

            const point = pointAt(curve, root);
            if (!point.exists()) {
                return { point: nowhere(), parameter };
            }

            return { point, parameter: root };
        },

        /** Whether the crossing is on the other figure itself, not only on its line or circle (a graph has no ends) */
        isOnOther(first, second, point) {
            const curve = findCurve(first, second);
            if (curve == null) {
                return false;
            }

            const otherFigure = other(curve, first, second);
            return otherFigure.isFunctionGraph === true || otherFigure.hitTest(point) != null;
        },
    };
})();
