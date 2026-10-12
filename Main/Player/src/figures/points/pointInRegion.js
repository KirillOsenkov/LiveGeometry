// Port of Main/Avalonia/DynamicGeometry/Figures/Points/PointInRegion.cs: a point kept inside
// a figure with an inside (isRegion: a polygon, a circle, a sector, a path). Its place is an
// offset from the figure's center, worked out from the figure as it is now; where that falls
// outside, the point sits at the nearest point inside and the offset keeps the wished-for
// place. Only a drag (or the file) sets the offset. Left out: the grid's Free point verb.

class PointInRegion extends FreePoint {
    constructor() {
        super();

        /** Where the point wants to be, from the figure's center */
        this.offset = new Point(0, 0);

        // a hand-written file may give X and Y alone: the offset is taken from them once
        // the figure is in the drawing
        this.hasOffset = false;
    }

    readXml(element) {
        super.readXml(element);
        this.hasOffset = element.getAttribute("OffsetX") != null;
        if (this.hasOffset) {
            this.offset = new Point(Xml.readDouble(element, "OffsetX"), Xml.readDouble(element, "OffsetY"));
        }
    }

    get region() {
        const figure = this.dependencies.length > 0 ? this.dependencies[0] : null;
        return figure != null && figure.isRegion === true ? figure : null;
    }

    capturePlace() {
        return this.offset;
    }

    restorePlace(place) {
        this.offset = place;
        this.recalculateAndUpdateVisual();
    }

    allowMove() {
        return !this.locked;
    }

    /** The point goes to the nearest place inside, and that is its place from now on */
    moveToCore(newPosition) {
        const region = this.region;
        if (region != null && region.exists && region.hasInside) {
            const inside = region.nearestPointInside(newPosition);
            const center = region.center;
            if (inside.exists() && center.exists()) {
                newPosition = inside;
                this.offset = inside.minus(center);
                this.hasOffset = true;
            }
        }

        super.moveToCore(newPosition);
    }

    recalculate() {
        const region = this.region;
        if (region == null || !allExist(this.dependencies) || !region.hasInside) {
            this.exists = false;
            return;
        }

        const center = region.center;
        if (!center.exists()) {
            this.exists = false;
            return;
        }

        if (!this.hasOffset) {
            this.offset = this.coordinates.minus(center);
            this.hasOffset = true;
        }

        const place = region.nearestPointInside(center.plus(this.offset));
        if (!place.exists()) {
            this.exists = false;
            return;
        }

        this.coordinates = place;
        this.exists = true;
    }
}

FigureTypes.register("PointInRegion", PointInRegion);
