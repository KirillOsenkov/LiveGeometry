using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Media.Imaging;
using DynamicGeometry;

namespace LiveGeometry;

/// <summary>
/// "LiveGeometry.Desktop.exe --check &lt;folder&gt; &lt;output folder&gt;": opens every drawing under
/// the folder in the real editor, one after another, writes a PNG of each next to a report of
/// what went wrong (load errors, figures that don't exist) and exits. For triaging a pile of
/// old files without clicking through them.
/// </summary>
public partial class MainView
{
    public static string CheckFolder { get; set; }

    public static string CheckOutputFolder { get; set; }

    public static string ModernizeFolder { get; set; }

    readonly List<string> checkMessages = new List<string>();

    async void RunCheck(string folder, string outputFolder)
    {
        Directory.CreateDirectory(outputFolder);
        DrawingHost.DrawingControl.DrawingAttach += drawing => drawing.Status += text =>
        {
            if (!string.IsNullOrEmpty(text) && !text.StartsWith("Processed in"))
            {
                checkMessages.Add(text);
            }
        };
        DrawingHost.UnhandledException += (s, e) => checkMessages.Add("EXCEPTION " + e.Exception.Message);
        MessageBox.Handler = text => checkMessages.Add("MESSAGE " + text);

        // written as it goes: a file that takes the process down is then the last one named
        var reportPath = Path.Combine(outputFolder, "report.txt");
        File.WriteAllText(reportPath, "");
        var report = new StringBuilder();
        var files = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories)
            .Where(f => f.EndsWith(".lgf", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".dgf", StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f)
            .ToArray();
        ShowEditor();

        // as the gallery would show them: without the tools on a small screen
        ribbonChoice = !IsSmallScreen;
        UpdateRibbon();
        foreach (var file in files)
        {
            var relative = Path.GetRelativePath(folder, file);
            checkMessages.Clear();
            File.AppendAllText(reportPath, report.ToString());
            report.Clear();
            File.AppendAllText(reportPath, "=== " + relative + Environment.NewLine);
            try
            {
                OpenDrawing(Path.GetFileName(file), File.ReadAllBytes(file));
                await Task.Delay(50);
                var drawing = DrawingHost.CurrentDrawing;
                if (!GalleryDrawing.HasCaption(drawing))
                {
                    // a captioned drawing was laid out for the window on opening
                    drawing.CoordinateSystem.ZoomExtend();
                }
                await Task.Delay(120);

                var figures = drawing.Figures.Where(f => !(f is CartesianGrid)).ToArray();
                // the figures that fail on their own (their dependencies are fine): the
                // rest just inherit that
                var missing = figures
                    .Where(f => !f.Exists && f.Dependencies.All(d => d.Exists))
                    .Select(f => f.GetType().Name + " " + f.Name + Describe(f) + " <- " + string.Join(",", f.Dependencies.Select(d => d.Name)))
                    .ToArray();
                var counts = figures.GroupBy(f => f.GetType().Name).OrderByDescending(g => g.Count()).Select(g => g.Key + "x" + g.Count());
                report.AppendLine("  figures: " + figures.Length + " (" + string.Join(" ", counts) + ")");
                if (missing.Length > 0)
                {
                    report.AppendLine("  NOT EXISTING (" + figures.Count(f => !f.Exists) + " in all): " + string.Join("; ", missing));
                    report.AppendLine("  points: " + string.Join(" ", figures.OfType<IPoint>().Select(p =>
                        p.Name + (p.Exists ? "(" + p.Coordinates.X.ToString("0.##") + "," + p.Coordinates.Y.ToString("0.##") + ")" : "(-)"))));
                }

                foreach (var message in checkMessages)
                {
                    report.AppendLine("  " + message.Replace("\n", " | "));
                }

                var control = DrawingHost.DrawingControl;
                var size = new PixelSize((int)control.Bounds.Width, (int)control.Bounds.Height);
                using var bitmap = new RenderTargetBitmap(size, new Avalonia.Vector(96, 96));
                bitmap.Render(control);
                var outputName = relative.Replace('\\', '_').Replace('/', '_');
                bitmap.Save(Path.Combine(outputFolder, outputName + ".png"), new PngBitmapEncoderOptions());

                // and the drawing in today's format: how a .dgf gets converted
                File.WriteAllText(Path.Combine(outputFolder, Path.ChangeExtension(outputName, ".lgf")), drawing.SaveAsText());
            }
            catch (Exception ex)
            {
                report.AppendLine("  CRASH " + ex.Message);
            }
        }

        File.AppendAllText(reportPath, report.ToString());
        Console.WriteLine("checked " + files.Length + " files");
        Environment.Exit(0);
    }

    /// <summary>
    /// "--modernize &lt;folder&gt;": writes into a file what loading it upgrades, so that the
    /// file itself is current. Two upgrades so far, each a one-off for the gallery: drawings
    /// marked IntersectionOrder="Legacy" get the swapped Algorithm attributes of the
    /// intersections the old circle/line order got wrong, and the mark dropped (2026-09-22);
    /// drawings before Version 1 get the offsets of their point labels and measurements in
    /// pixels - at the zoom this window shows the drawing at, which for a captioned drawing
    /// is the gallery's fit, so the labels stay exactly where they were on the desktop
    /// (2026-09-23; the window should be a landscape one) - and Version="1". Nothing else in
    /// the file changes.
    /// </summary>
    async void RunModernize(string folder)
    {
        ShowEditor();
        int changedFiles = 0;
        foreach (var file in Directory.GetFiles(folder, "*.lgf"))
        {
            var document = XDocument.Load(file, LoadOptions.PreserveWhitespace);
            bool legacyOrder = (string)document.Root.Attribute("IntersectionOrder") == "Legacy";
            bool unitOffsets = document.Root.ReadDouble("Version") < 1;
            if (!legacyOrder && !unitOffsets)
            {
                continue;
            }

            OpenDrawing(Path.GetFileName(file), File.ReadAllBytes(file));
            await Task.Delay(50);
            var drawing = DrawingHost.CurrentDrawing;
            var figures = document.Root.Element("Figures");
            var changes = new List<string>();
            if (legacyOrder)
            {
                var byName = drawing.Figures.OfType<IntersectionPoint>().ToDictionary(p => p.Name);
                int swapped = 0;
                foreach (var element in figures.Elements("IntersectionPoint"))
                {
                    var point = byName[(string)element.Attribute("Name")];
                    if (point.AlgorithmName != (string)element.Attribute("Algorithm"))
                    {
                        element.SetAttributeValue("Algorithm", point.AlgorithmName);
                        swapped++;
                    }
                }

                document.Root.Attribute("IntersectionOrder").Remove();
                changes.Add(swapped + " intersections swapped");
            }

            if (unitOffsets)
            {
                double unitLength = drawing.CoordinateSystem.UnitLength;
                int converted = 0;
                foreach (var element in figures.Elements())
                {
                    if (drawing.Figures[(string)element.Attribute("Name")] is LabelWithOffset && element.Attribute("OffsetX") != null)
                    {
                        var x = element.ReadDouble("OffsetX") * unitLength;
                        var y = -element.ReadDouble("OffsetY") * unitLength;
                        element.SetAttributeValue("OffsetX", System.Math.Round(x, 1).ToStringInvariant());
                        element.SetAttributeValue("OffsetY", System.Math.Round(y, 1).ToStringInvariant());
                        converted++;
                    }
                }

                document.Root.SetAttributeValue("Version", "1");
                changes.Add(converted + " label offsets to pixels at " + unitLength.ToString("0.#") + " px/unit");
            }

            var settings = new XmlWriterSettings() { Indent = true, Encoding = new UTF8Encoding(false), NewLineChars = "\r\n" };
            using (var writer = XmlWriter.Create(file, settings))
            {
                document.Save(writer);
            }

            Console.WriteLine(Path.GetFileName(file) + ": " + string.Join(", ", changes));
            changedFiles++;
        }

        Console.WriteLine("modernized " + changedFiles + " files");
        Environment.Exit(0);
    }

    public static string RecaptionFolder { get; set; }

    // explanations whose single line breaks are meant: a numbered list, a formula on a line of its own
    static readonly HashSet<string> structuredCaptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "BestFitCircle.lgf", "Ceva.lgf", "ParabolaGraph.lgf", "Pentagon.lgf", "Pythagoras.lgf", "SteinersProblem.lgf"
    };

    static readonly Regex formulaLine = new Regex(@"^(\d+\.\s|\[|\S+ = |\S+ \+ \S+ = )");

    /// <summary>
    /// "--recaption &lt;folder&gt;": a one-off that moved the gallery captions onto the screen
    /// (2026-09-23). The Title and Description labels of every drawing get the pin, offsets
    /// and column width that <see cref="GalleryDrawing.Fit"/> gives them in this window
    /// (which should be a landscape one), instead of X and Y, become clickable, and the
    /// explanations lose the line breaks that wrapped them by hand for a fixed width - the
    /// column wraps them now. Nothing else in the file changes.
    /// </summary>
    async void RunRecaption(string folder)
    {
        ShowEditor();
        int changedFiles = 0;
        foreach (var file in Directory.GetFiles(folder, "*.lgf"))
        {
            var document = XDocument.Load(file, LoadOptions.PreserveWhitespace);
            var labels = document.Root.Element("Figures").Elements("Label")
                .Where(e => (string)e.Attribute("Name") == GalleryDrawing.TitleName || (string)e.Attribute("Name") == GalleryDrawing.DescriptionName)
                .ToArray();
            if (labels.Length != 2)
            {
                continue;
            }

            var fileName = Path.GetFileName(file);
            var descriptionElement = labels.First(e => (string)e.Attribute("Name") == GalleryDrawing.DescriptionName);
            descriptionElement.SetAttributeValue("Text", Unwrap((string)descriptionElement.Attribute("Text"), structuredCaptions.Contains(fileName)));

            // laid out here, then written back
            using (var stream = new MemoryStream())
            {
                document.Save(stream);
                OpenDrawing(fileName, stream.ToArray());
            }

            await Task.Delay(50);
            var drawing = DrawingHost.CurrentDrawing;
            foreach (var element in labels)
            {
                var label = (Label)drawing.Figures[(string)element.Attribute("Name")];
                element.Attribute("X")?.Remove();
                element.Attribute("Y")?.Remove();

                // clickable again: a pinned label is dragged by its offset, which is how a
                // caption too long for a phone is pulled up to be read
                element.Attribute("IsHitTestVisible")?.Remove();
                element.SetAttributeValue("Pin", label.Pin.ToString());
                element.SetAttributeValue("OffsetX", label.PinOffset.X.ToStringInvariant());
                element.SetAttributeValue("OffsetY", label.PinOffset.Y.ToStringInvariant());
                element.SetAttributeValue("WrapWidth", label.WrapWidth.ToStringInvariant());
                element.SetAttributeValue("Backdrop", "true");
            }

            var settings = new XmlWriterSettings() { Indent = true, Encoding = new UTF8Encoding(false), NewLineChars = "\r\n" };
            using (var writer = XmlWriter.Create(file, settings))
            {
                document.Save(writer);
            }

            changedFiles++;
        }

        Console.WriteLine("recaptioned " + changedFiles + " files");
        Environment.Exit(0);
    }

    /// <summary>
    /// Joins the lines of a paragraph with spaces; a blank line still separates paragraphs.
    /// In a structured text a line after one ending with a colon, a numbered item and a
    /// formula (something = ...) stay on their own line.
    /// </summary>
    static string Unwrap(string text, bool structured)
    {
        var lines = text.Split(new[] { @"\n" }, StringSplitOptions.None);
        var sb = new StringBuilder(lines[0]);
        for (int i = 1; i < lines.Length; i++)
        {
            var line = lines[i];
            var previous = lines[i - 1];
            bool hardBreak = line.Length == 0
                || previous.Length == 0
                || (structured && (previous.EndsWith(":") || formulaLine.IsMatch(line)));
            sb.Append(hardBreak ? @"\n" : " ");
            sb.Append(line);
        }

        return sb.ToString();
    }

    static string Describe(IFigure figure)
    {
        if (figure is PointByCoordinates point)
        {
            return " X=\"" + point.XExpression.Text + "\" Y=\"" + point.YExpression.Text + "\""
                + (point.XExpression.IsValid ? "" : " (X invalid)")
                + (point.YExpression.IsValid ? "" : " (Y invalid)");
        }

        return "";
    }
}
