using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.TextFormatting;
using Avalonia.Threading;

namespace DynamicGeometry;

/// <summary>
/// One character (an emoji) in <see cref="EmojiFont"/>, its ink centered in the control and
/// about as tall as it: a point's character (<see cref="PointMarker"/>), and the swatches of
/// the Emoji tab. Until the font is loaded it draws in a fallback font, then again.
/// </summary>
public class EmojiGlyph : Control
{
    static readonly IBrush SelectionPlate = new SolidColorBrush(Color.FromArgb(70, 0, 120, 215));

    string text;
    public string Text
    {
        get => text;
        set
        {
            text = value;
            InvalidateVisual();
        }
    }

    IBrush foreground;

    /// <summary>For a character that isn't a color emoji (★); black if null</summary>
    public IBrush Foreground
    {
        get => foreground;
        set
        {
            foreground = value;
            InvalidateVisual();
        }
    }

    bool isHighlighted;

    /// <summary>A plate behind the character: its point is selected</summary>
    public bool IsHighlighted
    {
        get => isHighlighted;
        set
        {
            isHighlighted = value;
            InvalidateVisual();
        }
    }

    /// <summary>Room around the character, on each side</summary>
    public double Inset { get; set; }

    protected override Size MeasureOverride(Size availableSize)
    {
        return new Size(double.IsNaN(Width) ? 0 : Width, double.IsNaN(Height) ? 0 : Height);
    }

    public override void Render(DrawingContext context)
    {
        if (string.IsNullOrEmpty(Text))
        {
            return;
        }

        var size = Bounds.Size;
        if (IsHighlighted)
        {
            double radius = System.Math.Max(size.Width, size.Height) * 0.62;
            context.DrawEllipse(SelectionPlate, pen: null, new Point(size.Width / 2, size.Height / 2), radius, radius);
        }

        DrawCharacter(context, Text, new Rect(size).Deflate(Inset), Foreground);
        if (!EmojiFont.IsLoaded)
        {
            EmojiFont.EnsureLoaded().ContinueWith(
                _ => Dispatcher.UIThread.Post(InvalidateVisual),
                System.Threading.Tasks.TaskScheduler.Default);
        }
    }

    /// <summary>
    /// Draws the character so that its ink (not its line box, which has room above and below)
    /// is centered in the box and about as tall as it
    /// </summary>
    static void DrawCharacter(DrawingContext context, string character, Rect box, IBrush foreground)
    {
        double fontSize = System.Math.Max(1, System.Math.Min(box.Width, box.Height));
        using var layout = new TextLayout(
            character,
            new Typeface(EmojiFont.Family),
            fontSize,
            foreground ?? Brushes.Black);
        var ink = GetInkBounds(layout);
        var origin = new Point(
            box.Center.X - ink.Center.X,
            box.Center.Y - ink.Center.Y);
        layout.Draw(context, origin);
    }

    static Rect GetInkBounds(TextLayout layout)
    {
        var whole = new Rect(0, 0, layout.WidthIncludingTrailingWhitespace, layout.Height);
        if (layout.TextLines.Count != 1)
        {
            return whole;
        }

        var line = layout.TextLines[0];
        var bottom = line.Height + line.OverhangAfter;
        var left = line.Start + line.OverhangLeading;
        var right = line.Start + line.WidthIncludingTrailingWhitespace - line.OverhangTrailing;
        var ink = new Rect(left, bottom - line.Extent, right - left, line.Extent);
        if (!(ink.Width > 0 && ink.Height > 0))
        {
            return whole;
        }

        return ink;
    }
}
