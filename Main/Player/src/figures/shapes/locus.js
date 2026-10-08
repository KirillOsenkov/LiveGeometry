// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/Locus.cs: the curve a point traces as
// another point slides along its figure, sampled adaptively.

class Locus extends Curve {
    /** Even steps of the sliding point that adaptive sampling starts from */
    static InitialSteps = 60;

    /** In pixels: how far the curve may stray from the straight piece drawn between two samples */
    static Tolerance = 0.5;

    /** In pixels: two samples still this far apart when halving the step can't bring them closer are a jump */
    static JumpPixels = 20;

    static MaxDepth = 12;
    static MaxSamples = 1000;

    /** Steps outward past an open end of the sliding point's figure, each twice as long as the last */
    static OpenEndSteps = 10;

    /** In sizes of the window: a traced point further than this from it is taken for a gap */
    static FarReach = 20;

    constructor() {
        super();
        this.figuresToRecalculate = null;

        /** A fixed number of even steps of the sliding point, or 0 for as many as it takes */
        this.samples = 0;
        this.sampleParameters = [];
        this.samplePoints = [];
    }

    get isLocus() {
        return true;
    }

    /** What lies between the point that slides and the point traced, worked out anew each time */
    recalculate() {
        const point = this.dependencies.length > 0 && this.dependencies[0].isPoint === true ? this.dependencies[0] : null;
        const pointOnFigure = this.dependencies.length > 1 && this.dependencies[1] instanceof PointOnFigure ? this.dependencies[1] : null;
        this.figuresToRecalculate = point != null && pointOnFigure != null && point !== pointOnFigure
            ? DependencyAlgorithms.findImpactedDependencyChain(pointOnFigure, point)
            : null;
    }

    onDependenciesChanged() {
        this.figuresToRecalculate = null;
    }

    get hasChain() {
        return this.figuresToRecalculate != null && this.figuresToRecalculate.length > 0;
    }

    /** Where the traced point is when the sliding point is at the parameter; the caller puts the sliding point back */
    trace(sliding, traced, parameter) {
        sliding.parameter = parameter;
        for (const figure of this.figuresToRecalculate) {
            figure.updateExistence();
            figure.recalculate();
        }

        return traced.exists && traced.coordinates.exists() ? traced.coordinates : Curve.Gap;
    }

    getPoints(result) {
        this.sampleParameters = [];
        this.samplePoints = [];
        if (!this.hasChain) {
            return;
        }

        const point = this.dependencies[0];
        const pointOnFigure = this.dependencies[1];
        const domain = pointOnFigure.linearFigure.getParameterDomain();
        const oldParameter = pointOnFigure.parameter;
        if (this.samples > 0) {
            const steps = this.samples;
            for (let i = 0; i <= steps; i++) {
                const lambda = i === steps ? domain[1] : domain[0] + (domain[1] - domain[0]) * i / steps;
                this.sampleParameters.push(lambda);
                this.samplePoints.push(this.trace(pointOnFigure, point, lambda));
            }
        } else {
            this.sampleAdaptively(pointOnFigure, point, domain);
        }

        result.push(...this.samplePoints);
        this.trace(pointOnFigure, point, oldParameter);
    }

    readXml(element) {
        super.readXml(element);
        this.samples = Math.trunc(Math.max(0, Math.min(Locus.MaxSamples, Xml.readDouble(element, "Samples"))));
    }

    /**
     * InitialSteps even steps, then, round after round, every step halved where the curve
     * strays from the straight piece between two samples by more than Tolerance pixels, or
     * a gap begins or ends there - at most MaxDepth rounds and MaxSamples samples. Past the
     * open ends of a line, a ray or a graph the steps go outward, each twice as long.
     */
    sampleAdaptively(sliding, traced, domain) {
        const coordinateSystem = this.drawing.coordinateSystem;
        const minX = coordinateSystem.minimalVisibleX;
        const maxX = coordinateSystem.maximalVisibleX;
        const minY = coordinateSystem.minimalVisibleY;
        const maxY = coordinateSystem.maximalVisibleY;
        const windowCenter = new Point((minX + maxX) / 2, (minY + maxY) / 2);
        const reach = Locus.FarReach * Math.max(maxX - minX, maxY - minY);
        const unitLength = coordinateSystem.unitLength;

        let samplesLeft = Locus.MaxSamples;
        const take = parameter => {
            samplesLeft--;
            const place = this.trace(sliding, traced, parameter);
            return {
                parameter,
                point: place.exists() && place.distance(windowCenter) <= reach ? place : Curve.Gap,
                done: false
            };
        };

        let samples = [];
        const openEnds = Locus.getOpenEnds(sliding.linearFigure);
        const span = domain[1] - domain[0];
        if (openEnds.low) {
            for (let i = Locus.OpenEndSteps - 1; i >= 0; i--) {
                samples.push(take(domain[0] - span * Math.pow(2, i)));
            }
        }

        for (let i = 0; i <= Locus.InitialSteps; i++) {
            samples.push(take(i === Locus.InitialSteps ? domain[1] : domain[0] + span * i / Locus.InitialSteps));
        }

        if (openEnds.high) {
            for (let i = 0; i < Locus.OpenEndSteps; i++) {
                samples.push(take(domain[1] + span * Math.pow(2, i)));
            }
        }

        // a round halves every piece that isn't done
        let round = 0;
        let halved = true;
        while (halved && round < Locus.MaxDepth && samplesLeft > 0) {
            round++;
            halved = false;
            const next = [];
            for (let i = 0; i < samples.length; i++) {
                const from = samples[i];
                if (from.done || i === samples.length - 1 || samplesLeft <= 0) {
                    next.push(from);
                    continue;
                }

                const to = samples[i + 1];
                const fromExists = from.point.exists();
                const toExists = to.point.exists();
                if (!fromExists && !toExists) {
                    from.done = true;
                    next.push(from);
                    continue;
                }

                const middle = take((from.parameter + to.parameter) / 2);
                if (fromExists && toExists && middle.point.exists()
                    && Locus.distanceToSegment(middle.point, from.point, to.point) * unitLength <= Locus.Tolerance) {
                    from.done = true;
                    next.push(from);
                    continue;
                }

                next.push(from);
                next.push(middle);
                halved = true;
            }

            samples = next;
        }

        // pieces still far apart after every round of halving are jumps: not joined, unless the samples ran out first
        const ranOut = samplesLeft <= 0 && halved;
        for (let i = 0; i < samples.length; i++) {
            const sample = samples[i];
            this.sampleParameters.push(sample.parameter);
            this.samplePoints.push(sample.point);
            if (!ranOut && !sample.done && i < samples.length - 1) {
                const to = samples[i + 1];
                if (sample.point.exists() && to.point.exists() && sample.point.distance(to.point) * unitLength > Locus.JumpPixels) {
                    this.sampleParameters.push((sample.parameter + to.parameter) / 2);
                    this.samplePoints.push(Curve.Gap);
                }
            }
        }
    }

    static distanceToSegment(point, from, to) {
        const dx = to.x - from.x;
        const dy = to.y - from.y;
        const lengthSquared = dx * dx + dy * dy;
        let ratio = lengthSquared > 0 ? ((point.x - from.x) * dx + (point.y - from.y) * dy) / lengthSquared : 0;
        ratio = Math.max(0, Math.min(1, ratio));
        return point.distance(new Point(from.x + dx * ratio, from.y + dy * ratio));
    }

    /** Which ends of the parameters of the figure are only where the window cuts it: { low, high } */
    static getOpenEnds(figure) {
        if (figure.isFunctionGraph === true) {
            return { low: true, high: true };
        }

        if (figure instanceof LineBase && !(figure instanceof Segment)) {
            return { high: true, low: !(figure instanceof Ray) || (figure instanceof AngleBisector && figure.wholeLine) };
        }

        if (figure instanceof Locus && figure.dependencies.length > 1 && figure.dependencies[1] instanceof PointOnFigure) {
            return Locus.getOpenEnds(figure.dependencies[1].linearFigure);
        }

        return { low: false, high: false };
    }

    // A point on the locus keeps the sliding point's parameter on its own figure

    getNearestParameterFromPoint(point) {
        if (this.samplePoints.length === 0) {
            return 0;
        }

        let nearest = Number.MAX_VALUE;
        let parameter = this.sampleParameters[0];
        for (let i = 0; i < this.samplePoints.length - 1; i++) {
            const from = this.samplePoints[i];
            const to = this.samplePoints[i + 1];
            const dx = to.x - from.x;
            const dy = to.y - from.y;
            const lengthSquared = dx * dx + dy * dy;
            let ratio = lengthSquared > 0 ? ((point.x - from.x) * dx + (point.y - from.y) * dy) / lengthSquared : 0;
            ratio = Math.max(0, Math.min(1, ratio));
            const x = from.x + dx * ratio - point.x;
            const y = from.y + dy * ratio - point.y;
            const distance = x * x + y * y;
            if (distance < nearest) {
                nearest = distance;
                parameter = this.sampleParameters[i] + (this.sampleParameters[i + 1] - this.sampleParameters[i]) * ratio;
            }
        }

        return parameter;
    }

    /** Worked out, not read off the drawn curve: the sliding point is put at the parameter and back */
    getPointFromParameter(parameter) {
        const point = this.dependencies.length > 0 && this.dependencies[0].isPoint === true ? this.dependencies[0] : null;
        const pointOnFigure = this.dependencies.length > 1 && this.dependencies[1] instanceof PointOnFigure ? this.dependencies[1] : null;
        if (point == null || pointOnFigure == null || !this.hasChain || !isValidValue(parameter)) {
            return Point.infinite;
        }

        const oldParameter = pointOnFigure.parameter;
        const result = this.trace(pointOnFigure, point, parameter);
        this.trace(pointOnFigure, point, oldParameter);
        return result.exists() ? result : Point.infinite;
    }

    getParameterDomain() {
        const pointOnFigure = this.dependencies.length > 1 && this.dependencies[1] instanceof PointOnFigure ? this.dependencies[1] : null;
        return pointOnFigure != null ? pointOnFigure.linearFigure.getParameterDomain() : super.getParameterDomain();
    }
}

FigureTypes.register("Locus", Locus);
