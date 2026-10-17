// Port of Main/Avalonia/LiveGeometry/Figures/Controls/NameLayout.cs: names that keep clear.
// A point's name, a figure's name or a text attached to a figure is drawn where its offset
// puts it unless something covers that place - another name, a measurement, an angle's
// number or another point - and then at the nearest free place around its anchor, not
// across the lines through it, until its own place is free again. Only where it is drawn
// changes; a new one goes to the first free place.

class NameLayout {
    /** In square pixels: an overlap this small is a touch, not a cover */
    static OverlapAllowance = 2;

    /** In degrees: the steps around the anchor a name tries, the nearest first */
    static StepDegrees = 15;

    static Steps = 12;

    /** In pixels: kept between a name and a point */
    static PointGap = 2;

    /** In pixels: kept between a name that has moved and the other names, so that A and B side by side don't read as AB */
    static Spacing = 6;

    /** In pixels: how far into a name at its own place another name may reach, both ways, before it moves */
    static HomeDepth = 2;

    static DefaultPointSize = 8;

    /** In pixels: how near the anchor a line must run to be one through it */
    static OnAnchor = 1.5;

    static OverlapCost = 0.5;
    static CrossingCost = 0.5;

    // the names of a drawing drawn away from their place
    static nudged = new WeakMap();

    static rechecking = false;

    static namesOf(drawing) {
        let names = NameLayout.nudged.get(drawing);
        if (names == null) {
            names = new Set();
            NameLayout.nudged.set(drawing, names);
        }

        return names;
    }

    /** Where the name is drawn, its top-left corner in pixels, given its place there */
    static place(label, place) {
        return NameLayout.choose(label, place, false);
    }

    /** Where a new name goes, given its default place: the nearest free place to it, off the lines through the anchor */
    static firstPlace(label, place) {
        return NameLayout.choose(label, place, true);
    }

    static choose(label, place, isFirst) {
        const drawing = label.drawing;
        const size = label.measureSize();
        if (!(size.width > 0 && size.height > 0) || !place.exists()) {
            NameLayout.forget(label, drawing);
            return place;
        }

        const coordinateSystem = drawing.coordinateSystem;
        const anchor = coordinateSystem.toPhysical(label.anchor);
        const home = Rect.fromPoints(place, size);
        if (!anchor.exists()) {
            NameLayout.forget(label, drawing);
            return place;
        }

        const distance = home.center.distance(anchor);
        const reach = distance + 2 * Math.max(size.width, size.height) + NameLayout.DefaultPointSize;
        const ownPoint = label.ownPoint;
        let ownMarker = null;
        const points = [];
        const names = [];
        for (const figure of drawing.figures.list) {
            if (figure instanceof LabelWithOffset) {
                if (figure !== label) {
                    const name = figure.shownBox();
                    if (name != null && NameLayout.distanceToBox(anchor, name) <= reach) {
                        names.push(name);
                    }
                }
            } else if (figure.isPoint === true) {
                const marker = NameLayout.markerBox(figure, coordinateSystem);
                if (marker != null && NameLayout.distanceToBox(anchor, marker) <= reach) {
                    if (figure === ownPoint) {
                        ownMarker = marker;
                    } else {
                        points.push(marker);
                    }
                }
            }
        }

        // a place it goes to keeps Spacing from the other names
        const obstacles = points.concat(names.map(name => NameLayout.inflate(name, NameLayout.Spacing / 2)));

        // a place the user left is theirs, lines across it or not; a new one keeps off them
        let lines = null;
        const crossing = box => {
            if (lines == null) {
                lines = [];
                for (const figure of new Set(label.figuresThroughAnchor())) {
                    NameLayout.addLine(figure, coordinateSystem, anchor, reach, lines);
                }
            }

            const inner = NameLayout.inflate(box, -1);
            return lines.reduce((sum, line) => sum + NameLayout.clippedLength(line.from, line.to, inner), 0);
        };

        // a name's own place is covered by a point over it, or by a name that overlaps it by more than HomeDepth both ways
        const homeOverlap = NameLayout.overlap(home, points) + names.reduce((sum, name) => sum + NameLayout.deep(home, name), 0);
        const homeCrossing = isFirst && homeOverlap <= NameLayout.OverlapAllowance ? crossing(home) : 0;
        if (homeOverlap <= NameLayout.OverlapAllowance && homeCrossing === 0) {
            NameLayout.forget(label, drawing);
            return place;
        }

        const cost = (box, steps) => {
            const overlap = NameLayout.overlap(box, obstacles) + (ownMarker != null ? NameLayout.intersectionArea(box, ownMarker) : 0);
            const across = crossing(box);
            return {
                isFree: overlap <= NameLayout.OverlapAllowance && across === 0,
                value: steps + overlap * NameLayout.OverlapCost + across * NameLayout.CrossingCost
            };
        };

        // where it went last time, while that is still free
        if (!isFirst && label.nudge != null) {
            const kept = NameLayout.centered(anchor.plus(label.nudge), size);
            if (cost(kept, 0).isFree) {
                return kept.topLeft;
            }
        }

        // around the anchor, the nearest first: at the place's distance, then a ring further out
        let direction = home.center.minus(anchor);
        if (direction.length() < 1) {
            direction = new Point(0, -1);
        }

        direction = direction.scale(1 / direction.length());
        const ringStep = Math.max(size.width, size.height) / 2 + 2 * NameLayout.PointGap;
        let best = home;
        let bestCost = homeOverlap * NameLayout.OverlapCost + homeCrossing * NameLayout.CrossingCost;
        for (let ring = 0; ring < 2; ring++) {
            for (let step = 1; step <= NameLayout.Steps; step++) {
                for (const sign of [1, -1]) {
                    const turned = NameLayout.rotate(direction, sign * step * NameLayout.StepDegrees);
                    const box = NameLayout.placeOut(anchor, turned, distance + ring * ringStep, size, ownMarker);
                    const result = cost(box, step + ring * NameLayout.Steps);
                    if (result.isFree) {
                        return isFirst ? box.topLeft : NameLayout.remember(label, drawing, anchor, box);
                    }

                    if (result.value < bestCost) {
                        bestCost = result.value;
                        best = box;
                    }
                }
            }
        }

        // nowhere free: the place that covers least, which may be its own
        if (isFirst) {
            return best.topLeft;
        }

        if (best === home) {
            NameLayout.forget(label, drawing);
            return place;
        }

        return NameLayout.remember(label, drawing, anchor, best);
    }

    static remember(label, drawing, anchor, box) {
        label.nudge = box.center.minus(anchor);
        NameLayout.namesOf(drawing).add(label);
        return box.topLeft;
    }

    /** The name is at its own place again, or has left the drawing */
    static forget(label, drawing) {
        label.nudge = null;
        if (drawing != null) {
            NameLayout.nudged.get(drawing)?.delete(label);
        }
    }

    /** A label was drawn somewhere else: the names drawn away from their place look again */
    static recheck(moved) {
        const drawing = moved.drawing;
        const names = drawing != null ? NameLayout.nudged.get(drawing) : null;
        if (NameLayout.rechecking || names == null || names.size === 0) {
            return;
        }

        NameLayout.rechecking = true;
        try {
            for (const name of [...names]) {
                if (name !== moved && name.drawing === drawing) {
                    name.updateVisual();
                }
            }
        } finally {
            NameLayout.rechecking = false;
        }
    }

    /** A shown point's marker on the canvas, in pixels, with PointGap around it */
    static markerBox(point, coordinateSystem) {
        if (!point.visible || !point.exists) {
            return null;
        }

        const center = coordinateSystem.toPhysical(point.coordinates);
        if (!center.exists()) {
            return null;
        }

        const size = point.pointSize > 0 ? point.pointSize : NameLayout.DefaultPointSize;
        return new Rect(
            center.x - size / 2 - NameLayout.PointGap,
            center.y - size / 2 - NameLayout.PointGap,
            size + 2 * NameLayout.PointGap,
            size + 2 * NameLayout.PointGap);
    }

    /** A line or circle that runs through the anchor, as pieces in pixels near it */
    static addLine(figure, coordinateSystem, anchor, reach, lines) {
        if (!figure.visible || !figure.exists) {
            return;
        }

        if (figure.isLine === true) {
            const a = coordinateSystem.toPhysical(figure.coordinates.p1);
            const b = coordinateSystem.toPhysical(figure.coordinates.p2);
            const along = b.minus(a);
            const length = along.length();
            if (!(length > 0) || !a.exists() || !b.exists()) {
                return;
            }

            const unit = along.scale(1 / length);
            const far = anchor.minus(a).length() + reach;
            let from = a;
            let to = b;
            if (figure instanceof Ray) {
                to = a.plus(unit.scale(far));
            } else if (!(figure instanceof Segment) && !(figure instanceof Vector)) {
                from = a.minus(unit.scale(far));
                to = a.plus(unit.scale(far));
            }

            if (NameLayout.distanceToSegment(anchor, from, to) <= NameLayout.OnAnchor) {
                lines.push({ from, to });
            }

            return;
        }

        if (figure.isCircle === true) {
            const center = coordinateSystem.toPhysical(figure.center);
            const radius = figure.radius * coordinateSystem.unitLength;
            if (!(radius > 0) || !center.exists() || Math.abs(anchor.minus(center).length() - radius) > NameLayout.OnAnchor) {
                return;
            }

            const chords = 180;
            let previous = center.plus(new Point(radius, 0));
            for (let i = 1; i <= chords; i++) {
                const angle = 2 * Math.PI * i / chords;
                const next = center.plus(new Point(radius * Math.cos(angle), radius * Math.sin(angle)));
                if (NameLayout.distanceToSegment(anchor, previous, next) <= reach) {
                    lines.push({ from: previous, to: next });
                }

                previous = next;
            }
        }
    }

    /** A box of the size centered on the direction from the anchor at the distance, pushed out until it clears the anchor's own point */
    static placeOut(anchor, direction, distance, size, ownMarker) {
        let box = NameLayout.centered(anchor.plus(direction.scale(distance)), size);
        if (ownMarker == null) {
            return box;
        }

        const limit = distance + Math.max(size.width, size.height);
        while (NameLayout.intersects(box, ownMarker) && distance < limit) {
            distance += 1;
            box = NameLayout.centered(anchor.plus(direction.scale(distance)), size);
        }

        return box;
    }

    static centered(center, size) {
        return new Rect(center.x - size.width / 2, center.y - size.height / 2, size.width, size.height);
    }

    static rotate(vector, degrees) {
        const radians = degrees * Math.PI / 180;
        const cos = Math.cos(radians);
        const sin = Math.sin(radians);
        return new Point(vector.x * cos - vector.y * sin, vector.x * sin + vector.y * cos);
    }

    static inflate(box, by) {
        return new Rect(box.x - by, box.y - by, box.width + 2 * by, box.height + 2 * by);
    }

    static intersects(a, b) {
        return a.x < b.right && b.x < a.right && a.y < b.bottom && b.y < a.bottom;
    }

    static intersectionArea(a, b) {
        const width = Math.min(a.right, b.right) - Math.max(a.x, b.x);
        const height = Math.min(a.bottom, b.bottom) - Math.max(a.y, b.y);
        return width > 0 && height > 0 ? width * height : 0;
    }

    /** The area of an overlap deeper than HomeDepth both ways; none for a shallower one */
    static deep(a, b) {
        const width = Math.min(a.right, b.right) - Math.max(a.x, b.x);
        const height = Math.min(a.bottom, b.bottom) - Math.max(a.y, b.y);
        return width > NameLayout.HomeDepth && height > NameLayout.HomeDepth ? width * height : 0;
    }

    static overlap(box, obstacles) {
        let sum = 0;
        for (const obstacle of obstacles) {
            sum += NameLayout.intersectionArea(box, obstacle);
        }

        return sum;
    }

    static distanceToBox(point, box) {
        const dx = Math.max(Math.max(box.x - point.x, point.x - box.right), 0);
        const dy = Math.max(Math.max(box.y - point.y, point.y - box.bottom), 0);
        return Math.sqrt(dx * dx + dy * dy);
    }

    static distanceToSegment(point, from, to) {
        const along = to.minus(from);
        const squared = along.x * along.x + along.y * along.y;
        let t = squared > 0 ? ((point.x - from.x) * along.x + (point.y - from.y) * along.y) / squared : 0;
        t = Math.min(Math.max(t, 0), 1);
        return point.minus(from.plus(along.scale(t))).length();
    }

    /** How much of the piece lies in the box, in pixels (Liang-Barsky) */
    static clippedLength(from, to, box) {
        const dx = to.x - from.x;
        const dy = to.y - from.y;
        const range = { low: 0, high: 1 };
        if (!NameLayout.clip(-dx, from.x - box.x, range)
            || !NameLayout.clip(dx, box.right - from.x, range)
            || !NameLayout.clip(-dy, from.y - box.y, range)
            || !NameLayout.clip(dy, box.bottom - from.y, range)) {
            return 0;
        }

        return (range.high - range.low) * Math.sqrt(dx * dx + dy * dy);
    }

    static clip(p, q, range) {
        if (p === 0) {
            return q >= 0;
        }

        const r = q / p;
        if (p < 0) {
            if (r > range.high) {
                return false;
            }

            range.low = Math.max(range.low, r);
        } else {
            if (r < range.low) {
                return false;
            }

            range.high = Math.min(range.high, r);
        }

        return true;
    }
}
