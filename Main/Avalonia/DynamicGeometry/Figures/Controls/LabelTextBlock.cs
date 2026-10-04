using System;
using Avalonia;
using Avalonia.Controls;

namespace DynamicGeometry;

/// <summary>
/// The text of a label. Avalonia's TextBlock lays its text out again when it is arranged, in
/// the arranged size, and leaves out every line that doesn't fit whole into its height. On an
/// iPhone (and nowhere else) the gallery's explanation of the Circumscribed Circle ended at
/// "When does the", also when dragged up into view: the arranged layout needed a line more
/// than the measured one. Laid out at least as high as it measured, it still lost the line.
/// So the text keeps the layout it was measured with, at the width it was measured at and
/// with no height limit: no line is ever left out, and one too many would stick out of the
/// box, not vanish.
/// </summary>
public class LabelTextBlock : TextBlock
{
    double measuredWidth = double.PositiveInfinity;
    double measuredHeight = double.PositiveInfinity;

    // styles that target TextBlock apply to this one too
    protected override Type StyleKeyOverride => typeof(TextBlock);

    protected override Size MeasureOverride(Size availableSize)
    {
        measuredWidth = availableSize.Width;
        measuredHeight = availableSize.Height;
        return base.MeasureOverride(availableSize);
    }

    /// <summary>
    /// Measured with no height limit (on the canvas, always), the text was laid out then just
    /// as it is wanted here, and TextBlock's own arrange would throw that layout away and make
    /// the same one again - every time the text changes. A caption with a live number in it
    /// was laid out twice at every move of a drag, a third of the time the move took.
    /// </summary>
    protected override Size ArrangeOverride(Size finalSize)
    {
        if (double.IsPositiveInfinity(measuredHeight))
        {
            return finalSize;
        }

        double width = double.IsInfinity(measuredWidth) ? finalSize.Width : measuredWidth;
        base.ArrangeOverride(new Size(width, double.PositiveInfinity));
        return finalSize;
    }
}
