using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Skia.Helpers;
using SkiaSharp;

namespace LiveGeometry;

/// <summary>
/// A picture of a visual as it is on screen right now, at its own size: the canvas of the
/// drawing, for saving and for the clipboard. Only what the visual itself draws: the side
/// panel and the status bar lie over the canvas, they are not in it.
/// </summary>
public static class ViewportImage
{
    /// <param name="scaling">
    /// Pixels per unit of layout: at the scaling of the screen the picture has the pixels
    /// that are on the screen
    /// </param>
    public static byte[] ToPng(Visual visual, double scaling)
    {
        var size = PixelSize.FromSize(visual.Bounds.Size, scaling);
        using var bitmap = new RenderTargetBitmap(size, new Vector(96 * scaling, 96 * scaling));
        bitmap.Render(visual);
        using var stream = new MemoryStream();
        bitmap.Save(stream, new PngBitmapEncoderOptions());
        return stream.ToArray();
    }

    /// <summary>
    /// The same drawing calls that paint the screen, written down as SVG by Skia: lines and
    /// shapes are paths, and so is text (<see cref="SvgTextOutlines"/>), so that the picture
    /// needs no font. In units of layout, which is what a pixel is to SVG.
    /// </summary>
    public static async Task<byte[]> ToSvg(Visual visual)
    {
        var size = visual.Bounds.Size;
        using var drawn = new MemoryStream();

        // the document is complete only when the canvas is gone
        using (var canvas = SKSvgCanvas.Create(SKRect.Create((float)size.Width, (float)size.Height), drawn))
        {
            await DrawingContextHelper.RenderAsync(
                canvas,
                visual,
                new Rect(size),
                new Vector(96, 96));
        }

        // as Skia wrote it, white space and all: some of it is inside the text elements
        drawn.Position = 0;
        var document = XDocument.Load(drawn, LoadOptions.PreserveWhitespace);
        SvgTextOutlines.Convert(document);

        // Skia gives the picture a size but no box, and only with one does it scale when
        // something shows it at another size
        var root = document.Root;
        if (root.Attribute("viewBox") == null)
        {
            root.SetAttributeValue("viewBox", "0 0 " + (string)root.Attribute("width") + " " + (string)root.Attribute("height"));
        }

        using var stream = new MemoryStream();
        var settings = new XmlWriterSettings() { Encoding = new UTF8Encoding(false) };
        using (var writer = XmlWriter.Create(stream, settings))
        {
            document.Save(writer);
        }

        return stream.ToArray();
    }
}
