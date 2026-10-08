// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/BezierPathSmoother.cs: the automatic
// handles of a Bezier path worked out from its anchors. A handle the user has set (an
// offset, a point of the drawing) stays as it is and the automatic ones fit themselves to
// it. Every method gives the same curve, only bigger, smaller, turned or mirrored, for the
// anchors so moved. A pure function: nothing here touches a figure.

/** How a Bezier path works out the handles left to it */
const BezierPathSmoothing = {
    /** On their anchors: the path goes from anchor to anchor in straight lines */
    None: "None",

    /** Hobby's (METAFONT's): the curvature changes as little as it can across each anchor; four anchors on a circle give a circle */
    Hobby: "Hobby",

    /** Centripetal Catmull-Rom: each anchor's handles along the line through its neighbors, from those alone */
    CatmullRom: "CatmullRom",

    /** The natural cubic spline (chord length): the curvature doesn't jump at all, which the whole path takes part in */
    NaturalSpline: "NaturalSpline"
};

/** The anchors, the pieces between them and what is known at each anchor */
class SmootherAnchors {
    /** points: Point[]; ins, outs: (Point | null)[] - null for a handle to work out */
    constructor(points, ins, outs, closed) {
        this.points = points;
        this.ins = ins;
        this.outs = outs;
        this.closed = closed;
        const count = points.length;
        this.pieceCount = count < 2 ? 0 : closed ? count : count - 1;
        this.chords = new Array(this.pieceCount);
        this.lengths = new Array(this.pieceCount);
        const size = count === 0 ? 0 : Math.max(...points.map(p => Math.max(Math.abs(p.x), Math.abs(p.y))));
        const tiny = 1e-12 * (1 + size);
        for (let i = 0; i < this.pieceCount; i++) {
            this.chords[i] = points[(i + 1) % count].minus(points[i]);
            this.lengths[i] = BezierPathSmoother.length(this.chords[i]);
            if (!(this.lengths[i] > tiny)) {
                this.lengths[i] = 0;
            }
        }

        this.resultIns = new Array(count);
        this.resultOuts = new Array(count);
        for (let i = 0; i < count; i++) {
            this.resultIns[i] = ins[i] ?? new Point(0, 0);
            this.resultOuts[i] = outs[i] ?? new Point(0, 0);
        }
    }

    get count() {
        return this.points.length;
    }

    /** The piece that comes to the anchor; -1 at the start of an open path */
    pieceBefore(anchor) {
        return this.closed ? (anchor + this.count - 1) % this.count : anchor - 1;
    }

    /** The piece that leaves the anchor; -1 at the end of an open path */
    pieceAfter(anchor) {
        return anchor < this.pieceCount ? anchor : -1;
    }

    next(anchor) {
        return (anchor + 1) % this.count;
    }

    hasLength(piece) {
        return piece >= 0 && this.lengths[piece] > 0;
    }

    /** Whether both handles of the anchor are automatic and both pieces there have a length: the angle there is worked out */
    isSmooth(anchor) {
        return this.ins[anchor] == null
            && this.outs[anchor] == null
            && this.hasLength(this.pieceBefore(anchor))
            && this.hasLength(this.pieceAfter(anchor));
    }

    /** Where the path leaves the anchor, when that is given: by its out handle, or for an automatic one, straight on from the in handle. Null: free. */
    outDirection(anchor) {
        if (this.outs[anchor] != null) {
            return BezierPathSmoother.unit(this.outs[anchor]);
        }

        if (this.pieceBefore(anchor) >= 0 && this.ins[anchor] != null) {
            return BezierPathSmoother.unit(this.ins[anchor].negate());
        }

        return null;
    }

    /** Where the path comes to the anchor going, when that is given */
    inDirection(anchor) {
        if (this.ins[anchor] != null) {
            return BezierPathSmoother.unit(this.ins[anchor].negate());
        }

        if (this.pieceAfter(anchor) >= 0 && this.outs[anchor] != null) {
            return BezierPathSmoother.unit(this.outs[anchor]);
        }

        return null;
    }

    /** The stretches of pieces between anchors where the angle is given or free, each with the smooth anchors inside it; none when every anchor of a closed path is smooth */
    runs() {
        const runs = [];
        for (let start = 0; start < this.count; start++) {
            if (this.isSmooth(start) || !this.hasLength(this.pieceAfter(start))) {
                continue;
            }

            // the anchors of the run, the first and the last not smooth
            const run = [start];
            let anchor = start;
            do {
                anchor = this.next(anchor);
                run.push(anchor);
            }
            while (this.isSmooth(anchor));
            runs.push(run);
        }

        return runs;
    }

    get allSmooth() {
        if (!this.closed) {
            return false;
        }

        for (let i = 0; i < this.count; i++) {
            if (!this.isSmooth(i)) {
                return false;
            }
        }

        return true;
    }

    /** How far the path turns at the anchor, from the chord before to the chord after */
    turn(anchor) {
        return BezierPathSmoother.angle(this.chords[this.pieceBefore(anchor)], this.chords[this.pieceAfter(anchor)]);
    }
}

const BezierPathSmoother = {
    /**
     * The offsets of the handles from their anchors: those given (not null) as they are, the
     * automatic ones (null) worked out. Returns { ins, outs }.
     */
    smooth(anchors, ins, outs, closed, smoothing, tension) {
        const path = new SmootherAnchors(anchors, ins, outs, closed);
        if (smoothing !== BezierPathSmoothing.None && path.pieceCount > 0 && tension > 0 && isValidValue(tension)) {
            switch (smoothing) {
                case BezierPathSmoothing.Hobby:
                    BezierPathSmoother.hobby(path, tension);
                    break;
                case BezierPathSmoothing.CatmullRom:
                    BezierPathSmoother.catmullRom(path, tension);
                    break;
                case BezierPathSmoothing.NaturalSpline:
                    BezierPathSmoother.naturalSpline(path, tension);
                    break;
            }
        }

        return { ins: path.resultIns, outs: path.resultOuts };
    },

    // Hobby: the angles a piece leaves its first anchor and comes to its second at, from its
    // chord: theta counterclockwise, phi clockwise. At a smooth anchor the path doesn't turn:
    // phi there + theta there = -(the turn of the chords). Hobby's "mock curvature" is the
    // same on both sides of each smooth anchor, a linear equation per anchor. A free end has
    // the curvature of the piece's other end ("curl" 1).
    Curl: 1,

    // below this tension the equations lose their footing (MetaPost asks for 3/4 at least)
    LeastTension: 0.75,

    hobby(path, tension) {
        const c = 3 * Math.max(tension, BezierPathSmoother.LeastTension) - 1;
        const theta = new Array(path.pieceCount).fill(0);
        const phi = new Array(path.pieceCount).fill(0);
        if (path.allSmooth) {
            BezierPathSmoother.hobbyCycle(path, c, theta, phi);
        } else {
            for (const run of path.runs()) {
                BezierPathSmoother.hobbyRun(path, run, c, theta, phi);
            }
        }

        for (let piece = 0; piece < path.pieceCount; piece++) {
            if (!path.hasLength(piece)) {
                continue;
            }

            const next = path.next(piece);
            const chord = path.chords[piece];
            if (path.outs[piece] == null) {
                path.resultOuts[piece] = BezierPathSmoother.rotate(chord, theta[piece]).scale(BezierPathSmoother.velocity(theta[piece], phi[piece]) / (3 * tension));
            }

            if (path.ins[next] == null) {
                path.resultIns[next] = BezierPathSmoother.rotate(chord, -phi[piece]).scale(-BezierPathSmoother.velocity(phi[piece], theta[piece]) / (3 * tension));
            }
        }
    },

    /** A closed path smooth at every anchor: an equation per anchor, round and round */
    hobbyCycle(path, c, theta, phi) {
        const count = path.count;
        const sub = new Array(count).fill(0);
        const diagonal = new Array(count).fill(0);
        const superDiagonal = new Array(count).fill(0);
        const right = new Array(count).fill(0);
        for (let k = 0; k < count; k++) {
            const before = path.pieceBefore(k);
            const next = path.next(k);
            const dBefore = path.lengths[before];
            const dAfter = path.lengths[k];
            sub[k] = 1 / dBefore;
            diagonal[k] = c * (1 / dBefore + 1 / dAfter);
            superDiagonal[k] = 1 / dAfter;
            right[k] = -c * path.turn(k) / dBefore - path.turn(next) / dAfter;
        }

        const x = BezierPathSmoother.solveCyclic(sub, diagonal, superDiagonal, right);
        if (x == null) {
            return;
        }

        for (let k = 0; k < count; k++) {
            const next = path.next(k);
            theta[k] = x[k];
            phi[k] = -path.turn(next) - x[next];
        }
    },

    /** The pieces from one anchor that isn't smooth to the next: the angles at the smooth anchors between them, and at the two ends, given or free */
    hobbyRun(path, run, c, theta, phi) {
        // the unknowns: theta at the start of each piece, then phi at the end of the last
        const m = run.length - 1;
        const pieces = run.slice(0, m);
        const d = j => path.lengths[pieces[j]];
        const psi = j => path.turn(run[j]);
        const curl = BezierPathSmoother.Curl;

        const start = path.outDirection(run[0]);
        const end = path.inDirection(run[m]);
        if (m === 1 && start == null && end == null) {
            // two free ends: a straight piece
            theta[pieces[0]] = 0;
            phi[pieces[0]] = 0;
            return;
        }

        const sub = new Array(m + 1).fill(0);
        const diagonal = new Array(m + 1).fill(0);
        const superDiagonal = new Array(m + 1).fill(0);
        const right = new Array(m + 1).fill(0);
        if (start != null) {
            diagonal[0] = 1;
            right[0] = BezierPathSmoother.angle(path.chords[pieces[0]], start);
        } else {
            // curl: theta0 (c + curl) = phi1 (1 + curl c), phi1 = -psi1 - theta1 unless it is the end
            diagonal[0] = c + curl;
            if (m === 1) {
                superDiagonal[0] = -(1 + curl * c);
            } else {
                superDiagonal[0] = 1 + curl * c;
                right[0] = -(1 + curl * c) * psi(1);
            }
        }

        for (let j = 1; j < m; j++) {
            sub[j] = 1 / d(j - 1);
            diagonal[j] = c * (1 / d(j - 1) + 1 / d(j));
            right[j] = -c * psi(j) / d(j - 1);
            if (j + 1 < m) {
                superDiagonal[j] = 1 / d(j);
                right[j] -= psi(j + 1) / d(j);
            } else {
                // the next unknown is phi at the end itself
                superDiagonal[j] = -1 / d(j);
            }
        }

        if (end != null) {
            diagonal[m] = 1;
            right[m] = BezierPathSmoother.angle(end, path.chords[pieces[m - 1]]);
        } else {
            // curl: phi (c + curl) = theta of the last piece (1 + curl c)
            sub[m] = -(1 + curl * c);
            diagonal[m] = c + curl;
        }

        const x = BezierPathSmoother.solveTridiagonal(sub, diagonal, superDiagonal, right);
        if (x == null) {
            return;
        }

        for (let j = 0; j < m; j++) {
            theta[pieces[j]] = x[j];
            phi[pieces[j]] = j + 1 < m ? -psi(j + 1) - x[j + 1] : x[m];
        }
    },

    /** How long Hobby makes a handle, in thirds of the chord, for a piece that leaves at theta and comes in at phi: 1 for a straight piece, 1.1716 for a quarter of a circle */
    velocity(theta, phi) {
        const sinTheta = Math.sin(theta);
        const cosTheta = Math.cos(theta);
        const sinPhi = Math.sin(phi);
        const cosPhi = Math.cos(phi);
        const sqrt5 = Math.sqrt(5);
        const numerator = 2 + Math.sqrt(2) * (sinTheta - sinPhi / 16) * (sinPhi - sinTheta / 16) * (cosTheta - cosPhi);
        const denominator = 1 + (sqrt5 - 1) / 2 * cosTheta + (3 - sqrt5) / 2 * cosPhi;

        // (METAFONT's limit: a piece that turns almost all the way round)
        const longest = 4;
        return denominator > numerator / longest ? numerator / denominator : longest;
    },

    // Catmull-Rom: the tangent at an anchor from its two neighbors, the knots spaced by the
    // square roots of the chords (centripetal), the handles a third of the tangent times the
    // piece's spacing. Where there is only one neighbor to go by the handle points halfway to
    // the handle at the other end of its piece.
    catmullRom(path, tension) {
        const count = path.count;
        const inKnown = new Array(count).fill(false);
        const outKnown = new Array(count).fill(false);
        for (let k = 0; k < count; k++) {
            inKnown[k] = path.ins[k] != null;
            outKnown[k] = path.outs[k] != null;
            const before = path.pieceBefore(k);
            const after = path.pieceAfter(k);
            if (!path.hasLength(before) || !path.hasLength(after)) {
                continue;
            }

            const point = path.points[k];
            const previous = path.points[(k + count - 1) % count];
            const next = path.points[path.next(k)];
            const spacingBefore = Math.sqrt(path.lengths[before]);
            const spacingAfter = Math.sqrt(path.lengths[after]);
            const tangent = point.minus(previous).scale(1 / spacingBefore)
                .minus(next.minus(previous).scale(1 / (spacingBefore + spacingAfter)))
                .plus(next.minus(point).scale(1 / spacingAfter));
            const outOffset = tangent.scale(spacingAfter / 3);
            const inOffset = tangent.scale(-spacingBefore / 3);
            if (path.outs[k] == null && path.ins[k] == null) {
                path.resultOuts[k] = outOffset;
                path.resultIns[k] = inOffset;
                outKnown[k] = true;
                inKnown[k] = true;
            } else if (path.outs[k] == null && path.outDirection(k) != null) {
                path.resultOuts[k] = path.outDirection(k).scale(BezierPathSmoother.length(outOffset));
                outKnown[k] = true;
            } else if (path.ins[k] == null && path.inDirection(k) != null) {
                path.resultIns[k] = path.inDirection(k).scale(-BezierPathSmoother.length(inOffset));
                inKnown[k] = true;
            }
        }

        // the free ends, halfway to the handle at the other end of the piece (both free: a straight piece, the handles at its thirds)
        for (let piece = 0; piece < path.pieceCount; piece++) {
            if (!path.hasLength(piece)) {
                continue;
            }

            const first = piece;
            const second = path.next(piece);
            const from = path.points[first];
            const to = path.points[second];
            if (!outKnown[first] && !inKnown[second]) {
                path.resultOuts[first] = path.chords[piece].scale(1 / 3);
                path.resultIns[second] = path.chords[piece].scale(-1 / 3);
            } else if (!outKnown[first]) {
                path.resultOuts[first] = to.plus(path.resultIns[second]).minus(from).scale(0.5);
            } else if (!inKnown[second]) {
                path.resultIns[second] = from.plus(path.resultOuts[first]).minus(to).scale(0.5);
            }
        }

        BezierPathSmoother.scaleAutomatic(path, 1 / tension);
    },

    // Natural spline: the pieces as one curve with continuous first and second derivatives,
    // the parameter running along each piece as far as its chord is long: D at each anchor,
    // the out handle D d / 3, the in handle -D d / 3 with the chord before. A free end has no
    // curvature; a given direction is D itself.
    naturalSpline(path, tension) {
        if (path.allSmooth) {
            BezierPathSmoother.naturalCycle(path);
        } else {
            for (const run of path.runs()) {
                BezierPathSmoother.naturalRun(path, run);
            }
        }

        BezierPathSmoother.scaleAutomatic(path, 1 / tension);
    },

    naturalCycle(path) {
        const count = path.count;
        const sub = new Array(count).fill(0);
        const diagonal = new Array(count).fill(0);
        const superDiagonal = new Array(count).fill(0);
        const rightX = new Array(count).fill(0);
        const rightY = new Array(count).fill(0);
        for (let k = 0; k < count; k++) {
            const before = path.pieceBefore(k);
            const hBefore = path.lengths[before];
            const hAfter = path.lengths[k];
            sub[k] = 1 / hBefore;
            diagonal[k] = 2 * (1 / hBefore + 1 / hAfter);
            superDiagonal[k] = 1 / hAfter;
            const right = path.chords[before].scale(3 / (hBefore * hBefore)).plus(path.chords[k].scale(3 / (hAfter * hAfter)));
            rightX[k] = right.x;
            rightY[k] = right.y;
        }

        const x = BezierPathSmoother.solveCyclic(sub, diagonal, superDiagonal, rightX);
        const y = BezierPathSmoother.solveCyclic(sub, diagonal, superDiagonal, rightY);
        if (x == null || y == null) {
            return;
        }

        for (let k = 0; k < count; k++) {
            const derivative = new Point(x[k], y[k]);
            path.resultOuts[k] = derivative.scale(path.lengths[k] / 3);
            path.resultIns[k] = derivative.scale(-path.lengths[path.pieceBefore(k)] / 3);
        }
    },

    naturalRun(path, run) {
        const m = run.length - 1;
        const pieces = run.slice(0, m);
        const h = j => path.lengths[pieces[j]];
        const chord = j => path.chords[pieces[j]];

        const sub = new Array(m + 1).fill(0);
        const diagonal = new Array(m + 1).fill(0);
        const superDiagonal = new Array(m + 1).fill(0);
        const right = new Array(m + 1).fill(null);
        const start = path.outDirection(run[0]);
        if (start != null) {
            diagonal[0] = 1;
            right[0] = start;
        } else {
            diagonal[0] = 2;
            superDiagonal[0] = 1;
            right[0] = chord(0).scale(3 / h(0));
        }

        for (let j = 1; j < m; j++) {
            sub[j] = 1 / h(j - 1);
            diagonal[j] = 2 * (1 / h(j - 1) + 1 / h(j));
            superDiagonal[j] = 1 / h(j);
            right[j] = chord(j - 1).scale(3 / (h(j - 1) * h(j - 1))).plus(chord(j).scale(3 / (h(j) * h(j))));
        }

        const end = path.inDirection(run[m]);
        if (end != null) {
            diagonal[m] = 1;
            right[m] = end;
        } else {
            sub[m] = 1;
            diagonal[m] = 2;
            right[m] = chord(m - 1).scale(3 / h(m - 1));
        }

        const x = BezierPathSmoother.solveTridiagonal(sub, diagonal, superDiagonal, right.map(p => p.x));
        const y = BezierPathSmoother.solveTridiagonal(sub, diagonal, superDiagonal, right.map(p => p.y));
        if (x == null || y == null) {
            return;
        }

        for (let j = 0; j < m; j++) {
            const first = run[j];
            const second = run[j + 1];
            if (path.outs[first] == null) {
                path.resultOuts[first] = new Point(x[j], y[j]).scale(h(j) / 3);
            }

            if (path.ins[second] == null) {
                path.resultIns[second] = new Point(x[j + 1], y[j + 1]).scale(-h(j) / 3);
            }
        }
    },

    /** The automatic handles longer or shorter by the factor (the tension of the methods that have none of their own) */
    scaleAutomatic(path, factor) {
        for (let k = 0; k < path.count; k++) {
            if (path.ins[k] == null) {
                path.resultIns[k] = path.resultIns[k].scale(factor);
            }

            if (path.outs[k] == null) {
                path.resultOuts[k] = path.resultOuts[k].scale(factor);
            }
        }
    },

    length(vector) {
        return Math.sqrt(vector.x * vector.x + vector.y * vector.y);
    },

    unit(vector) {
        const length = BezierPathSmoother.length(vector);
        return length > 0 && isValidValue(length) ? vector.scale(1 / length) : null;
    },

    rotate(vector, angle) {
        const cos = Math.cos(angle);
        const sin = Math.sin(angle);
        return new Point(vector.x * cos - vector.y * sin, vector.x * sin + vector.y * cos);
    },

    /** The angle counterclockwise from one direction to the other, in (-pi, pi] */
    angle(from, to) {
        const angle = Math.atan2(from.x * to.y - from.y * to.x, from.x * to.x + from.y * to.y);
        return angle <= -Math.PI ? Math.PI : angle;
    },

    /** Thomas's algorithm: row i says sub[i] x[i-1] + diagonal[i] x[i] + super[i] x[i+1] = right[i]; null when it breaks down */
    solveTridiagonal(sub, diagonal, superDiagonal, right) {
        const count = diagonal.length;
        const superPrime = new Array(count).fill(0);
        const rightPrime = new Array(count).fill(0);
        for (let i = 0; i < count; i++) {
            const pivot = diagonal[i] - (i > 0 ? sub[i] * superPrime[i - 1] : 0);
            if (Math.abs(pivot) < 1e-300) {
                return null;
            }

            superPrime[i] = superDiagonal[i] / pivot;
            rightPrime[i] = (right[i] - (i > 0 ? sub[i] * rightPrime[i - 1] : 0)) / pivot;
        }

        const x = new Array(count).fill(0);
        for (let i = count - 1; i >= 0; i--) {
            x[i] = rightPrime[i] - (i + 1 < count ? superPrime[i] * x[i + 1] : 0);
        }

        return x.every(value => isValidValue(value)) ? x : null;
    },

    /** The same round a cycle: sub[0] goes with the last unknown and super[last] with the first (Sherman and Morrison's correction) */
    solveCyclic(sub, diagonal, superDiagonal, right) {
        const count = diagonal.length;
        if (count === 2) {
            const a = diagonal[0];
            const b = sub[0] + superDiagonal[0];
            const c = sub[1] + superDiagonal[1];
            const d = diagonal[1];
            const determinant = a * d - b * c;
            if (Math.abs(determinant) < 1e-300) {
                return null;
            }

            return [(right[0] * d - b * right[1]) / determinant, (a * right[1] - c * right[0]) / determinant];
        }

        const corner = superDiagonal[count - 1];
        const otherCorner = sub[0];
        const gamma = -diagonal[0];
        const changed = [...diagonal];
        changed[0] -= gamma;
        changed[count - 1] -= corner * otherCorner / gamma;
        const plainSub = [...sub];
        const plainSuper = [...superDiagonal];
        plainSub[0] = 0;
        plainSuper[count - 1] = 0;
        const x = BezierPathSmoother.solveTridiagonal(plainSub, changed, plainSuper, right);
        const u = new Array(count).fill(0);
        u[0] = gamma;
        u[count - 1] = corner;
        const z = BezierPathSmoother.solveTridiagonal(plainSub, changed, plainSuper, u);
        if (x == null || z == null) {
            return null;
        }

        const denominator = 1 + z[0] + otherCorner * z[count - 1] / gamma;
        if (Math.abs(denominator) < 1e-300) {
            return null;
        }

        const factor = (x[0] + otherCorner * x[count - 1] / gamma) / denominator;
        for (let i = 0; i < count; i++) {
            x[i] -= factor * z[i];
        }

        return x;
    }
};
