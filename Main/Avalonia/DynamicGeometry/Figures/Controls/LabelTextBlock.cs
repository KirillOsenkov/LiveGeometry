using System;
using Avalonia;
using Avalonia.Controls;

namespace DynamicGeometry;

/// <summary>
/// The text of a label. Avalonia's TextBlock lays its text out again when it is arranged, at
/// most as high as the space it is given, and leaves out every line that doesn't fit whole:
/// arranged a pixel shorter than it measured, a wrapped caption lost its last line. On a
/// phone the gallery's explanation of the Circumscribed Circle ended at "When does the",
/// also when dragged up into view. So the text is laid out at least as high as it measured.
/// </summary>
public class LabelTextBlock : TextBlock
{
    // styles that target TextBlock apply to this one too
    protected override Type StyleKeyOverride => typeof(TextBlock);

    protected override Size ArrangeOverride(Size finalSize)
    {
        return base.ArrangeOverride(new Size(finalSize.Width, System.Math.Max(finalSize.Height, DesiredSize.Height)));
    }
}
