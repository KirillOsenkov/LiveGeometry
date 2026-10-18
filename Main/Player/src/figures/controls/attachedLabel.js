// Port of Main/Avalonia/LiveGeometry/Figures/Controls/AttachedLabel.cs: a text attached to a
// figure, written next to a point, a segment, a line, a ray or a circle as its name would be
// (FigureAnchor) and moving with it, but saying whatever is typed. Plain text: no [...]
// parts. It keeps clear as a name does (NameLayout).

class AttachedLabel extends AnchoredLabel {
    readXml(element) {
        super.readXml(element);
        this.setText(Label.unescape(element.getAttribute("Text") ?? ""));
    }
}

FigureTypes.register("AttachedLabel", AttachedLabel);
