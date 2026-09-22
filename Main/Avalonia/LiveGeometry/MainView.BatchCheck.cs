using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
                var missing = figures.Where(f => !f.Exists).Select(f => f.GetType().Name + " " + f.Name).ToArray();
                var counts = figures.GroupBy(f => f.GetType().Name).OrderByDescending(g => g.Count()).Select(g => g.Key + "x" + g.Count());
                report.AppendLine("  figures: " + figures.Length + " (" + string.Join(" ", counts) + ")");
                if (missing.Length > 0)
                {
                    report.AppendLine("  NOT EXISTING: " + string.Join(", ", missing));
                }

                foreach (var message in checkMessages)
                {
                    report.AppendLine("  " + message.Replace("\n", " | "));
                }

                var control = DrawingHost.DrawingControl;
                var size = new PixelSize((int)control.Bounds.Width, (int)control.Bounds.Height);
                using var bitmap = new RenderTargetBitmap(size, new Avalonia.Vector(96, 96));
                bitmap.Render(control);
                var pngName = relative.Replace('\\', '_').Replace('/', '_') + ".png";
                bitmap.Save(Path.Combine(outputFolder, pngName), new PngBitmapEncoderOptions());
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
}
