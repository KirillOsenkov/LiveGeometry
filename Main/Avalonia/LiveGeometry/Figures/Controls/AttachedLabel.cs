using System.Xml;
using System.Xml.Linq;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// A text attached to a figure (2026-10-17): written next to a point, a segment, a line, a
/// ray or a circle as its name would be, or in the middle of a polygon
/// (<see cref="FigureAnchor"/>), and moving with it,
/// but saying whatever is typed - "r", "hypotenuse", "x = 4√2" - so that a name doesn't have
/// to be made into a comment. The Text tool makes one with a click on a figure; it is a
/// figure of its own (the Figure List has it, Delete takes it) and goes when its figure
/// goes. The text is plain: a free label's [...] parts make it depend on the figures they
/// name, and this one depends on its figure alone. It keeps clear as a name does
/// (<see cref="NameLayout"/>). The player has the twin (attachedLabel.js).
/// </summary>
public class AttachedLabel : AnchoredLabel
{
    /// <summary>"on segment AB": "Text “r” on segment AB"</summary>
    public override string Construction => Figure != null ? "on " + ConstructionText.Of(Figure) : null;

    /// <summary>
    /// Plain text, not the free label's box that puts references to figures in [...]: they
    /// would not be worked out here. A limitation of the dependency list, not of the idea:
    /// working out [...] parts (<see cref="LabelBase.ProcessText"/>) clears the label's
    /// dependencies and fills them with the figures the parts name, and this text's one
    /// dependency is the figure it hangs on. To allow them, that figure would have to stay
    /// first and the parts' figures come after it.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridMultiline]
    [PropertyGridDisallowMultiEdit]
    public override string Text
    {
        get => base.Text; set => base.Text = value;
    }

    /// <summary>No numbers in a plain text: no Decimals</summary>
    public override bool ShowsDecimals()
    {
        return false;
    }

    /// <summary>Puts the text with its top-left corner here, logical: from now on its place</summary>
    public void PlaceAt(Point topLeft)
    {
        Offset = ToPhysical(topLeft).Minus(ToPhysical(Anchor));
    }

    /// <summary>Lets go of the figure: a free text where this one is drawn (<see cref="TextAttachment.Detach"/>)</summary>
    [PropertyGridVisible]
    [ContextMenu]
    [PropertyGridName("Detach")]
    [PropertyGridIcon(PropertyGridIcon.Unlock)]
    public void Detach()
    {
        var free = TextAttachment.Detach(this);
        free.Drawing.RaiseDisplayProperties(free);
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);
        Text = Label.Unescape(element.ReadString("Text") ?? "");
    }

    public override void WriteXml(XmlWriter writer)
    {
        base.WriteXml(writer);
        writer.WriteAttributeString("Text", Label.Escape(Text));
    }
}
