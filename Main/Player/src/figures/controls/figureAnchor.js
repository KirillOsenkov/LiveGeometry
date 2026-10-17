// Port of Main/Avalonia/LiveGeometry/Figures/Controls/FigureAnchor.cs: where a label written
// next to a figure hangs (a figure's name, FigureLabel, and a text attached to a figure,
// AttachedLabel): an anchor on the figure, a first offset from it, and for a line, which
// runs past the window, a place kept on screen.

class FigureAnchor {
    /** How far in from the edge of the window a line's label sits, in pixels */
    static EdgeInset = 24;

    /** Between a segment and its label, in pixels */
    static SegmentGap = 6;

    /** The least room between a line's label and the edge of the window, in pixels */
    static EdgeMargin = 4;

    /** Where the offset is measured from, logical: a point itself, the middle of a segment, the upper left of a circle, near the edge of the window for a line or a ray */
    static of(figure, coordinateSystem) {
        if (figure != null && figure.isPoint === true) {
            return figure.coordinates;
        }

        if (figure instanceof Segment) {
            return figure.coordinates.midpoint;
        }

        if (figure instanceof Ray) {
            // the far end of what shows
            return FigureAnchor.inward(figure.onScreenCoordinates.p2, figure.onScreenCoordinates.p1, coordinateSystem);
        }

        if (figure instanceof LineBase) {
            // the end of the visible part nearer the top of the window
            const shown = figure.onScreenCoordinates;
            return shown.p1.y >= shown.p2.y
                ? FigureAnchor.inward(shown.p1, shown.p2, coordinateSystem)
                : FigureAnchor.inward(shown.p2, shown.p1, coordinateSystem);
        }

        if (figure != null && figure.isCircle === true) {
            // upper left, as GeoGebra puts it
            const center = figure.center;
            const radius = figure.radius * Math.sqrt(0.5);
            return new Point(center.x - radius, center.y + radius);
        }

        if (figure instanceof PolygonBase) {
            return FigureAnchor.centroid(figure.vertexCoordinates);
        }

        return figure != null ? figure.center : new Point();
    }

    /** The middle of a polygon's inside, weighted by area; the average of the vertices for one with no area */
    static centroid(vertices) {
        let area = 0;
        let x = 0;
        let y = 0;
        for (let i = 0; i < vertices.length; i++) {
            const a = vertices[i];
            const b = vertices[(i + 1) % vertices.length];
            const cross = a.x * b.y - b.x * a.y;
            area += cross;
            x += (a.x + b.x) * cross;
            y += (a.y + b.y) * cross;
        }

        if (Math.abs(area) < 1e-12 || vertices.length < 3) {
            return GeometryMath.midpoint(vertices);
        }

        return new Point(x / (3 * area), y / (3 * area));
    }

    /** The point EdgeInset pixels from the end towards the other end, logical */
    static inward(end, other, coordinateSystem) {
        if (coordinateSystem == null) {
            return end;
        }

        const from = coordinateSystem.toPhysical(end);
        const direction = RightAngleMark.direction(from, coordinateSystem.toPhysical(other));
        if (direction == null) {
            return end;
        }

        return coordinateSystem.toLogical(from.plus(direction.scale(FigureAnchor.EdgeInset)));
    }

    /** Where a new label goes, in pixels from the anchor: beside a segment or a line on its left, just inside a circle, to the right of a point */
    static defaultOffset(figure, size, coordinateSystem) {
        if (figure instanceof PolygonBase) {
            return new Point(-size.width / 2, -size.height / 2);
        }

        if (figure instanceof LineBase) {
            const coordinates = figure instanceof Segment ? figure.coordinates : figure.onScreenCoordinates;
            const direction = RightAngleMark.direction(coordinateSystem.toPhysical(coordinates.p1), coordinateSystem.toPhysical(coordinates.p2)) ?? new Point(1, 0);
            const normal = new Point(direction.y, -direction.x);
            const distance = FigureAnchor.SegmentGap + Math.abs(normal.x) * size.width / 2 + Math.abs(normal.y) * size.height / 2;
            return normal.scale(distance).minus(new Point(size.width / 2, size.height / 2));
        }

        if (figure.isCircle === true) {
            return new Point(4, 2);
        }

        if (figure.isPoint === true) {
            return new Point(figure.pointSize / 2 + PointLabel.Margin, -size.height / 2);
        }

        return new Point(6, 2);
    }

    /** Whether a line's anchor is near enough to the window to write its label */
    static isOnScreen(pixel, coordinateSystem) {
        const canvas = coordinateSystem.physicalSize;
        return pixel.exists()
            && pixel.x >= -FigureAnchor.EdgeInset && pixel.x <= canvas.x + FigureAnchor.EdgeInset
            && pixel.y >= -FigureAnchor.EdgeInset && pixel.y <= canvas.y + FigureAnchor.EdgeInset;
    }

    /** The top-left corner, in pixels, that brings a line's label back in, whole */
    static keptOnScreen(topLeft, size, coordinateSystem) {
        const canvas = coordinateSystem.physicalSize;
        return new Point(
            FigureAnchor.clamp(topLeft.x, FigureAnchor.EdgeMargin, canvas.x - size.width - FigureAnchor.EdgeMargin),
            FigureAnchor.clamp(topLeft.y, FigureAnchor.EdgeMargin, canvas.y - size.height - FigureAnchor.EdgeMargin));
    }

    static clamp(value, low, high) {
        return high < low ? low : Math.min(Math.max(value, low), high);
    }
}
