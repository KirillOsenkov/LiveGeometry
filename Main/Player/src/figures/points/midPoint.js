// Port of Main/Avalonia/DynamicGeometry/Figures/Points/MidPoint.cs

class MidPoint extends PointBase {
    recalculate() {
        this.coordinates = new Point(
            (this.point(0).x + this.point(1).x) / 2,
            (this.point(0).y + this.point(1).y) / 2);
    }
}

FigureTypes.register("MidPoint", MidPoint);
