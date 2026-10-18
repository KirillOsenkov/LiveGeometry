// Port of Main/Avalonia/LiveGeometry/Figures/Controls/FigureLabel.cs: the name of a line,
// ray, segment or circle, written next to it, at Offset pixels from an anchor on it.

class FigureLabel extends NameLabel {
    updateVisual() {
        if (this.figure != null) {
            this.setText(NameDisplay.format(this.figure.name));
        }

        super.updateVisual();
    }
}

FigureTypes.register("FigureLabel", FigureLabel);
