// Port of Main/Avalonia/LiveGeometry/Figures/Controls/AngleLabelLayout.cs: where the numbers
// of the angles at a vertex go while nobody has placed them by hand - each inside its own
// angle, past its mark, clear of the other numbers at the vertex and, where there is room,
// of the sides. A place is a turn off the bisector (a fraction of the half angle) and a push
// out past the mark; a number tries a few, and takes the cheapest. It remembers its place
// and starts from there at the next redraw: it slides by small steps, and goes further only
// when that is clearly cheaper (Hysteresis, Proximity), so it doesn't jump as figures move.

class AngleLabelLayout {
    /** In pixels: between a mark and its number */
    static Gap = 3;

    /** In pixels: kept between two numbers */
    static Spacing = 2;

    /** In pixels: how far a number is pushed out past where it clears its mark */
    static MaxPush = 90;

    /** In pixels: the steps of that push */
    static PushStep = 3;

    /** In pixels: what a number keeps clear of the vertex when its angle has no mark */
    static NoMarkClearance = 6;

    // in pixels of being pushed out: a square pixel of overlap, a pixel of a side through the
    // box, the middle of the box outside the angle, a turn all the way to a side
    static OverlapCost = 10;
    static SideCost = 1;
    static OutsideCost = 25;
    static TurnCost = 24;

    /** The farthest a number turns off the bisector, as a fraction of the half angle */
    static MaxTurn = 0.75;

    static Turns = [0, 0.25, -0.25, 0.5, -0.5, 0.75, -0.75];

    /** In pixels of being pushed out: what a place more than Near from where the number was must save to be taken */
    static Hysteresis = 15;

    /** In pixels of being pushed out, per pixel: what the way from where the number was costs */
    static Proximity = 1;

    /** In pixels: how far a number slides without the Hysteresis */
    static Near = 4;

    // the small steps a number slides by from where it was
    static SlideTurn = 0.02;
    static SlidePush = 1.5;

    /**
     * Puts the numbers of the angles at the vertex of this one that place themselves where
     * they go now: their offsets are set, and the others moved (the one asked about is
     * placed by its own redraw, which asks this)
     */
    static arrangeAt(number) {
        const drawing = number.drawing;
        const vertex = number.dependencies[0];
        const numbers = [...new Set(vertex.dependents)]
            .filter(other => other instanceof AngleMeasurement && (other === number || other.takesPartInLayout(vertex)))
            .sort((a, b) => drawing.figures.indexOf(a) - drawing.figures.indexOf(b));
        const placing = numbers.filter(other => other.placesItself);
        if (placing.length === 0) {
            return;
        }

        const corner = drawing.coordinateSystem.toPhysical(number.point(0));
        const obstacles = numbers
            .filter(other => !other.placesItself)
            .map(other => Rect.fromPoints(corner.plus(other.offset), other.measureSize()));
        const places = AngleLabelLayout.arrange(placing.map(other => other.layoutWedge()), obstacles, AngleLabelLayout.sidesOf(numbers));
        for (let i = 0; i < placing.length; i++) {
            placing[i].layoutPlacement = places[i].placement;
            placing[i].offset = places[i].topLeft.minus(corner);
            if (placing[i] !== number) {
                placing[i].moveToOffset();
            }
        }
    }

    /** The directions of the sides of the angles, out of their vertex, in radians in the plane */
    static sidesOf(numbers) {
        const sides = [];
        for (const number of numbers) {
            sides.push(GeometryMath.getAngle(number.point(0), number.point(1)));
            sides.push(GeometryMath.getAngle(number.point(0), number.point(2)));
        }

        return sides;
    }

    /**
     * Where the boxes of the wedges ({ vertex, start, sweep, isClockwise, clearance, size,
     * previous }: pixels but for the angles, which are in the plane) go, in their order -
     * { topLeft, placement } - kept clear of the obstacles (boxes that stay where they are)
     * and of the sides
     */
    static arrange(wedges, obstacles, sides) {
        // every number where it was, for the first pass to keep clear of
        const boxes = wedges.map(wedge => wedge.previous != null ? AngleLabelLayout.boxAtPlacement(wedge, wedge.previous) : null);
        const placements = new Array(wedges.length).fill(null);
        for (let pass = 0; pass < 2; pass++) {
            for (let i = 0; i < wedges.length; i++) {
                const others = obstacles.slice();
                for (let j = 0; j < boxes.length; j++) {
                    if (j !== i && boxes[j] != null) {
                        others.push(boxes[j]);
                    }
                }

                const placed = AngleLabelLayout.place(wedges[i], others, sides);
                boxes[i] = placed.box;
                placements[i] = placed.placement;
            }
        }

        return boxes.map((box, i) => ({ topLeft: box.topLeft, placement: placements[i] }));
    }

    static place(wedge, others, sides) {
        const bisector = AngleLabelLayout.bisector(wedge);
        const preferred = AngleLabelLayout.clearance(wedge, AngleLabelLayout.direction(bisector));
        const previousMiddle = wedge.previous != null ? AngleLabelLayout.boxAtPlacement(wedge, wedge.previous).center : null;
        let best = new Rect();
        let bestPlacement = { turn: 0, push: 0 };
        let bestCost = Infinity;

        // whether the place is clean: nothing over it, inside the angle
        const consider = (direction, clearance, placement) => {
            const distance = clearance + placement.push;
            const box = AngleLabelLayout.boxAt(wedge, direction, distance);
            const overlap = AngleLabelLayout.overlap(box, others);
            let across = 0;
            for (const side of sides) {
                across += AngleLabelLayout.chord(wedge.vertex, AngleLabelLayout.direction(side), box);
            }

            const inside = AngleLabelLayout.isMiddleInside(wedge, box);
            let cost = overlap * AngleLabelLayout.OverlapCost
                + across * AngleLabelLayout.SideCost
                + (inside ? 0 : AngleLabelLayout.OutsideCost)
                + Math.abs(placement.turn) * AngleLabelLayout.TurnCost
                + Math.max(0, distance - preferred);
            if (previousMiddle != null) {
                const way = box.center.distance(previousMiddle);
                cost += way * AngleLabelLayout.Proximity + (way > AngleLabelLayout.Near ? AngleLabelLayout.Hysteresis : 0);
            }

            if (cost < bestCost) {
                bestCost = cost;
                best = box;
                bestPlacement = placement;
            }

            return overlap === 0 && across === 0 && inside;
        };

        for (const turn of AngleLabelLayout.Turns) {
            const direction = AngleLabelLayout.direction(bisector + turn * wedge.sweep / 2);
            const clearance = AngleLabelLayout.clearance(wedge, direction);
            for (let push = 0; push <= AngleLabelLayout.MaxPush; push += AngleLabelLayout.PushStep) {
                // further out only costs more
                if (consider(direction, clearance, { turn, push })) {
                    break;
                }
            }
        }

        // where it was, and a small step in every way from there
        const from = wedge.previous;
        if (from != null) {
            for (const turnStep of [0, -AngleLabelLayout.SlideTurn, AngleLabelLayout.SlideTurn]) {
                const turn = Math.min(Math.max(from.turn + turnStep, -AngleLabelLayout.MaxTurn), AngleLabelLayout.MaxTurn);
                const direction = AngleLabelLayout.direction(bisector + turn * wedge.sweep / 2);
                const clearance = AngleLabelLayout.clearance(wedge, direction);
                for (const pushStep of [0, -AngleLabelLayout.SlidePush, AngleLabelLayout.SlidePush]) {
                    consider(direction, clearance, { turn, push: Math.min(Math.max(from.push + pushStep, 0), AngleLabelLayout.MaxPush) });
                }
            }
        }

        return { box: best, placement: bestPlacement };
    }

    static bisector(wedge) {
        return wedge.start + (wedge.isClockwise ? -wedge.sweep : wedge.sweep) / 2;
    }

    static boxAtPlacement(wedge, placement) {
        const direction = AngleLabelLayout.direction(AngleLabelLayout.bisector(wedge) + placement.turn * wedge.sweep / 2);
        return AngleLabelLayout.boxAt(wedge, direction, AngleLabelLayout.clearance(wedge, direction) + placement.push);
    }

    /** How far out along the direction the middle of the box goes at the least: where the box clears the mark */
    static clearance(wedge, direction) {
        const clearance = wedge.clearance + AngleLabelLayout.Gap;
        let low = 0;
        let high = clearance + wedge.size.width + wedge.size.height;
        for (let i = 0; i < 30; i++) {
            const middle = (low + high) / 2;
            if (AngleLabelLayout.distanceToBox(wedge.vertex, AngleLabelLayout.boxAt(wedge, direction, middle)) >= clearance) {
                high = middle;
            } else {
                low = middle;
            }
        }

        return high;
    }

    /** A unit vector in pixels (y down) along a direction in the plane (y up) */
    static direction(angle) {
        return new Point(Math.cos(angle), -Math.sin(angle));
    }

    static boxAt(wedge, direction, distance) {
        const middle = wedge.vertex.plus(direction.scale(distance));
        return new Rect(middle.x - wedge.size.width / 2, middle.y - wedge.size.height / 2, wedge.size.width, wedge.size.height);
    }

    /** Whether the middle of the box is inside the angle */
    static isMiddleInside(wedge, box) {
        const offset = box.center.minus(wedge.vertex);
        const angle = Math.atan2(-offset.y, offset.x);
        return AngleLabelLayout.normalize(wedge.isClockwise ? wedge.start - angle : angle - wedge.start) <= wedge.sweep;
    }

    static normalize(angle) {
        const turn = 2 * Math.PI;
        angle %= turn;
        return angle < 0 ? angle + turn : angle;
    }

    /** In pixels: how much of the ray from the origin along the direction (a unit vector) runs through the box */
    static chord(origin, direction, box) {
        const range = { near: 0, far: Infinity };
        if (!AngleLabelLayout.slab(origin.x, direction.x, box.x, box.right, range)
            || !AngleLabelLayout.slab(origin.y, direction.y, box.y, box.bottom, range)) {
            return 0;
        }

        return range.far - range.near;
    }

    static slab(origin, direction, low, high, range) {
        if (Math.abs(direction) < 1e-12) {
            return origin >= low && origin <= high;
        }

        const first = (low - origin) / direction;
        const second = (high - origin) / direction;
        range.near = Math.max(range.near, Math.min(first, second));
        range.far = Math.min(range.far, Math.max(first, second));
        return range.near <= range.far;
    }

    static distanceToBox(point, box) {
        const x = Math.max(Math.max(box.x - point.x, point.x - box.right), 0);
        const y = Math.max(Math.max(box.y - point.y, point.y - box.bottom), 0);
        return Math.sqrt(x * x + y * y);
    }

    /** The area the box overlaps the others by, each with Spacing around it */
    static overlap(box, others) {
        let sum = 0;
        for (const other of others) {
            const width = Math.min(box.right, other.right + AngleLabelLayout.Spacing) - Math.max(box.x, other.x - AngleLabelLayout.Spacing);
            const height = Math.min(box.bottom, other.bottom + AngleLabelLayout.Spacing) - Math.max(box.y, other.y - AngleLabelLayout.Spacing);
            if (width > 0 && height > 0) {
                sum += width * height;
            }
        }

        return sum;
    }
}
