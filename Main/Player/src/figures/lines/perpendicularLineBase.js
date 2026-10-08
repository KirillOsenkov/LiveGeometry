// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/PerpendicularLineBase.cs: a line that is
// perpendicular to another by construction, and says so with a RightAngleMark where they meet.

class PerpendicularLineBase extends LineTwoPoints {
    constructor() {
        super();
        this.rightAngleMark = null;
        this.cornerChosen = false;
    }

    get mark() {
        if (this.rightAngleMark == null) {
            this.rightAngleMark = new RightAngleMark(this);
        }

        return this.rightAngleMark;
    }

    /** { vertex, baseLine, pointAcross } where the right angle is, or null if it isn't there to be marked */
    tryGetRightAngle() {
        return null;
    }

    /** The figure drawn along the base line, if any */
    get baseFigure() {
        return null;
    }

    get showRightAngle() {
        return this.mark.isEnabled;
    }

    set showRightAngle(value) {
        this.mark.isEnabled = value;
    }

    get visible() {
        return super.visible;
    }

    set visible(value) {
        super.visible = value;
        if (!value) {
            this.mark.hide();
        } else if (this.drawing != null) {
            this.updateVisual();
        }
    }

    updateVisual() {
        super.updateVisual();
        if (!this.exists) {
            this.mark.hide();
            return;
        }

        const rightAngle = this.tryGetRightAngle();

        // once, when the line is first worked out; from then on the corner only changes by a click
        if (!this.cornerChosen) {
            this.cornerChosen = true;
            if (rightAngle != null) {
                this.mark.corner = RightAngleMark.getRoomiestCorner(rightAngle.vertex, rightAngle.baseLine, rightAngle.pointAcross);
            }
        }

        if (!this.visible || rightAngle == null || !rightAngle.hasRightAngle) {
            this.mark.hide();
            return;
        }

        this.mark.show(this.drawing, rightAngle.vertex, rightAngle.baseLine, this.baseFigure);
    }

    render(renderer) {
        super.render(renderer);
        if (this.isShown) {
            this.mark.render(renderer);
        }
    }

    readXml(element) {
        super.readXml(element);
        if (this.mark.readXml(element) != null) {
            this.cornerChosen = true;
        }
    }
}
