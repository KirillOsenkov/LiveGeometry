using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
                drawing.CoordinateSystem.ZoomExtend();
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
    /// "--modernize &lt;folder&gt;": a one-off for drawings marked IntersectionOrder="Legacy":
    /// loads each (which swaps the intersections the old circle/line order got wrong), writes
    /// the swapped Algorithm attributes into the file and drops the mark. Nothing else in the
    /// file changes.
    /// </summary>
    async void RunModernize(string folder)
    {
        ShowEditor();
        int changedFiles = 0;
        foreach (var file in Directory.GetFiles(folder, "*.lgf"))
        {
            var document = XDocument.Load(file, LoadOptions.PreserveWhitespace);
            if ((string)document.Root.Attribute("IntersectionOrder") != "Legacy")
            {
                continue;
            }

            OpenDrawing(Path.GetFileName(file), File.ReadAllBytes(file));
            await Task.Delay(50);
            var byName = DrawingHost.CurrentDrawing.Figures.OfType<IntersectionPoint>().ToDictionary(p => p.Name);
            int swapped = 0;
            foreach (var element in document.Root.Element("Figures").Elements("IntersectionPoint"))
            {
                var point = byName[(string)element.Attribute("Name")];
                if (point.AlgorithmName != (string)element.Attribute("Algorithm"))
                {
                    element.SetAttributeValue("Algorithm", point.AlgorithmName);
                    swapped++;
                }
            }

            document.Root.Attribute("IntersectionOrder").Remove();
            var settings = new XmlWriterSettings() { Indent = true, Encoding = new UTF8Encoding(false), NewLineChars = "\r\n" };
            using (var writer = XmlWriter.Create(file, settings))
            {
                document.Save(writer);
            }

            Console.WriteLine(Path.GetFileName(file) + ": " + swapped + " intersections swapped");
            changedFiles++;
        }

        Console.WriteLine("modernized " + changedFiles + " files");
        Environment.Exit(0);
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
