#:project ..\Main\Avalonia\LiveGeometry.Desktop\LiveGeometry.Desktop.csproj

// Checks the JavaScript player (Main/Player) against the library it is a port of.
//
//   dotnet run tools/playerparity.cs [-- [--structure] [--numbers] [--folder <drawings>] [--stop]]
//
// Structure: every figure kind the deserializer reads (DrawingDeserializer.FigureTypes) is
// registered in the player (FigureTypes.register) or listed in Main/Player/excluded.txt with
// a reason, and every function of the expression language (Functions) exists in
// src/expressions/functions.js. Numbers: every drawing of the folder (the gallery by
// default) loaded by both, dumped (Drawing.dump in JS, the same shape here: the figures in
// the list's order with the numbers that define them) and compared, then again after every
// free point is moved by a fixed offset. The player runs in headless Edge through
// tools/webauto.cs, on the dev page served by tools/serve.cs (both started here when they
// aren't running; --stop closes Edge at the end). Prints the differences and fails on any.
//
// What is not compared, by design: anything sized in pixels (the mark of an angle, label
// places), what is sampled by the window (loci, graphs: their points on them are compared),
// and the parts of composites (not in the list).

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using DynamicGeometry;
using LiveGeometry;

namespace LiveGeometryPlayerParity;

public class Program
{
    const int ServePort = 5005;
    const int EdgePort = 9333;
    const double Tolerance = 1e-6;
    static readonly Point dragOffset = new Point(0.3, 0.2);

    [STAThread]
    public static int Main(string[] args)
    {
        bool structure = args.Contains("--structure") || !args.Contains("--numbers");
        bool numbers = args.Contains("--numbers") || !args.Contains("--structure");
        bool stop = args.Contains("--stop");
        int folderIndex = Array.IndexOf(args, "--folder");
        var root = FindRepoRoot();
        var folder = folderIndex >= 0 && folderIndex + 1 < args.Length
            ? Path.GetFullPath(args[folderIndex + 1])
            : Path.Combine(root, "Main", "Avalonia", "LiveGeometry", "Gallery", "Drawings");

        App.UseInvariantCulture();
        AppBuilder.Configure<App>().UsePlatformDetect().WithInterFont().SetupWithoutStarting();
        Settings.Instance.AutoLabelPoints = false;

        int failures = 0;
        if (structure)
        {
            failures += CheckStructure(root);
        }

        if (numbers)
        {
            failures += CheckNumbers(root, folder, stop);
        }

        Console.WriteLine(failures == 0 ? "The player agrees with the library." : failures + " difference(s).");
        return failures == 0 ? 0 : 1;
    }

    static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(Directory.GetCurrentDirectory());
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "AGENTS.md")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName ?? throw new InvalidOperationException("Run from inside the repository.");
    }

    #region Structure

    static int CheckStructure(string root)
    {
        int failures = 0;
        var playerRoot = Path.Combine(root, "Main", "Player");
        // the kinds the player reads, and the classes it has (a base class is never in a
        // file, and has its twin without a registration)
        var registered = new HashSet<string>();
        var classes = new HashSet<string>();
        foreach (var file in Directory.GetFiles(Path.Combine(playerRoot, "src"), "*.js", SearchOption.AllDirectories))
        {
            var text = File.ReadAllText(file);
            foreach (Match match in Regex.Matches(text, @"FigureTypes\.register\(""(\w+)"""))
            {
                registered.Add(match.Groups[1].Value);
            }

            foreach (Match match in Regex.Matches(text, @"^class (\w+)", RegexOptions.Multiline))
            {
                classes.Add(match.Groups[1].Value);
            }
        }

        var excluded = new Dictionary<string, string>();
        var excludedPath = Path.Combine(playerRoot, "excluded.txt");
        foreach (var line in File.Exists(excludedPath) ? File.ReadAllLines(excludedPath) : Array.Empty<string>())
        {
            var trimmed = line.Trim();
            int colon = trimmed.IndexOf(':');
            if (trimmed.Length == 0 || trimmed.StartsWith("#") || colon < 0)
            {
                continue;
            }

            excluded[trimmed.Substring(0, colon).Trim()] = trimmed.Substring(colon + 1).Trim();
        }

        // what a file can name: a concrete kind the deserializer can make (the dictionary
        // holds every IFigure type, interfaces and bases included); a part made by its owner
        // (a polygon's vertex, a path's handle) has no such constructor
        var kinds = DrawingDeserializer.FigureTypes
            .Where(pair => !pair.Value.IsAbstract
                && !pair.Value.IsInterface
                && !pair.Value.IsGenericTypeDefinition
                && !pair.Value.IsNested
                && pair.Value.GetConstructor(Type.EmptyTypes) != null)
            .Select(pair => pair.Key)
            .OrderBy(k => k)
            .ToList();
        foreach (var kind in kinds)
        {
            if (!registered.Contains(kind) && !classes.Contains(kind) && !excluded.ContainsKey(kind))
            {
                Console.WriteLine("STRUCTURE: the library reads <" + kind + ">, the player has no figures/**/" + char.ToLowerInvariant(kind[0]) + kind.Substring(1) + ".js registering it (or list it in Main/Player/excluded.txt with a reason).");
                failures++;
            }
        }

        // (an AxisLine is made by the drawing, which the deserializer asks for it: no constructor of its own)
        foreach (var kind in registered.Where(kind => !DrawingDeserializer.FigureTypes.ContainsKey(kind)))
        {
            Console.WriteLine("STRUCTURE: the player registers <" + kind + ">, which the library doesn't have.");
            failures++;
        }

        foreach (var kind in excluded.Keys.Where(kind => registered.Contains(kind)))
        {
            Console.WriteLine("STRUCTURE: " + kind + " is in excluded.txt and registered in the player: take it off the list.");
            failures++;
        }

        // the functions of the expression language, by name
        var functionsText = File.ReadAllText(Path.Combine(playerRoot, "src", "expressions", "functions.js"));
        var playerFunctions = new HashSet<string>(Regex.Matches(functionsText, @"^    (\w+)\(", RegexOptions.Multiline).Select(m => m.Groups[1].Value));
        foreach (var method in typeof(Functions).GetMethods(BindingFlags.Public | BindingFlags.Static).Select(m => m.Name).Distinct().OrderBy(n => n))
        {
            if (!playerFunctions.Contains(method))
            {
                Console.WriteLine("STRUCTURE: Functions." + method + " has no twin in src/expressions/functions.js.");
                failures++;
            }
        }

        Console.WriteLine("structure: " + kinds.Count + " figure kinds, " + registered.Count + " in the player, " + excluded.Count + " excluded; " + playerFunctions.Count + " functions.");
        return failures;
    }

    #endregion

    #region Numbers

    static int CheckNumbers(string root, string folder, bool stop)
    {
        var files = Directory.GetFiles(folder, "*.lgf").OrderBy(f => f).ToList();
        var names = files.Select(Path.GetFileName).ToList();

        // the library's side, in this process
        var library = new Dictionary<string, (List<Dictionary<string, object>> Load, List<Dictionary<string, object>> Drag)>();
        foreach (var file in files)
        {
            var drawing = NewDrawing();
            drawing.AddFromXml(XElement.Parse(File.ReadAllText(file)));
            var load = Dump(drawing);
            Drag(drawing);
            library[Path.GetFileName(file)] = (load, Dump(drawing));
        }

        // the player's side, in headless Edge
        var relativeFolder = Path.GetRelativePath(Path.Combine(root, "Main"), folder).Replace('\\', '/');
        var player = DumpInPlayer(root, relativeFolder, names, stop);
        if (player == null)
        {
            return 1;
        }

        int failures = 0;
        foreach (var name in names)
        {
            if (!player.TryGetValue(name, out var dumps))
            {
                Console.WriteLine(name + ": the player gave no dump.");
                failures++;
                continue;
            }

            if (dumps.Error != null)
            {
                Console.WriteLine(name + ": the player threw: " + dumps.Error);
                failures++;
                continue;
            }

            var differences = new List<string>();
            Compare(library[name].Load, dumps.Load, "load", differences);
            Compare(library[name].Drag, dumps.Drag, "drag", differences);
            if (differences.Count > 0)
            {
                failures += differences.Count;
                Console.WriteLine(name + ": " + differences.Count + " difference(s)");
                foreach (var difference in differences.Take(12))
                {
                    Console.WriteLine("  " + difference);
                }
            }
        }

        Console.WriteLine("numbers: " + names.Count + " drawings, " + library.Values.Sum(d => d.Load.Count) + " figures compared twice.");
        return failures;
    }

    static Drawing NewDrawing()
    {
        var canvas = new Canvas { Width = 1000, Height = 700 };
        canvas.Measure(new Size(1000, 700));
        canvas.Arrange(new Rect(0, 0, 1000, 700));
        var drawing = new Drawing(canvas);
        drawing.UnhandledException += (_, arguments) => throw arguments.Exception;
        return drawing;
    }

    /// <summary>
    /// Every free point (not one on a figure) moved by the offset, and what is built on them
    /// worked out and drawn again as the Drag tool does it (Dragger, Actions.Move); the same
    /// in the player. (Drawing.Recalculate alone leaves the texts of labels behind.)
    /// </summary>
    static void Drag(Drawing drawing)
    {
        var points = drawing.Figures.Where(f => f.GetType() == typeof(FreePoint)).Cast<FreePoint>().ToList();
        foreach (var point in points)
        {
            point.MoveTo(point.Coordinates.Plus(dragOffset));
        }

        var dependents = DependencyAlgorithms.FindDescendants(f => f.Dependents, points.Cast<IFigure>().ToList());
        dependents.Reverse();
        foreach (var dependent in dependents)
        {
            dependent.RecalculateAndUpdateVisual();
        }
    }

    /// <summary>The same shape as Drawing.dump in the player</summary>
    static List<Dictionary<string, object>> Dump(Drawing drawing)
    {
        var result = new List<Dictionary<string, object>>();
        foreach (var figure in drawing.Figures)
        {
            if (figure is CartesianGrid)
            {
                continue;
            }

            var entry = new Dictionary<string, object>
            {
                ["name"] = figure.Name,
                ["kind"] = figure.GetType().Name,
                ["exists"] = figure.Exists,
                ["visible"] = figure.Visible
            };
            if (figure is IPoint point)
            {
                entry["coordinates"] = Numbers(point.Coordinates);
            }
            else if (figure is ILine line)
            {
                entry["p1"] = Numbers(line.Coordinates.P1);
                entry["p2"] = Numbers(line.Coordinates.P2);
            }
            else if (figure is IEllipse ellipse)
            {
                entry["center"] = Numbers(figure.Center);
                entry["semiMajor"] = Number(ellipse.SemiMajor);
                entry["semiMinor"] = Number(ellipse.SemiMinor);
            }
            else if (figure is IPolygonalChain chain)
            {
                entry["vertices"] = (chain.VertexCoordinates ?? Array.Empty<Point>()).Select(Numbers).ToList();
            }
            else if (figure is LabelBase label)
            {
                entry["text"] = label.ProcessedText;
            }
            else if (figure is INumber number)
            {
                entry["value"] = Number(number.Value);
            }

            result.Add(entry);
        }

        return result;
    }

    static object Number(double value)
    {
        return value.IsValidValue() ? System.Math.Round(value, 9) : value.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    static object[] Numbers(Point point)
    {
        return new[] { Number(point.X), Number(point.Y) };
    }

    static void Compare(List<Dictionary<string, object>> library, List<JsonElement> player, string stage, List<string> differences)
    {
        if (player == null)
        {
            differences.Add(stage + ": no dump from the player");
            return;
        }

        if (library.Count != player.Count)
        {
            differences.Add(stage + ": " + library.Count + " figures in the library, " + player.Count + " in the player");
        }

        int count = System.Math.Min(library.Count, player.Count);
        for (int i = 0; i < count; i++)
        {
            var expected = library[i];
            var actual = player[i];
            var name = (string)expected["name"];
            string where = stage + " " + expected["kind"] + " " + name;
            if (actual.GetProperty("name").GetString() != name)
            {
                differences.Add(where + ": the player's figure " + i + " is " + actual.GetProperty("name").GetString());
                continue;
            }

            foreach (var pair in expected)
            {
                if (!actual.TryGetProperty(pair.Key, out var value))
                {
                    differences.Add(where + ": the player's dump has no " + pair.Key);
                    continue;
                }

                // the mark of an angle is sized in pixels: its radius goes with the canvas
                bool pixels = (string)expected["kind"] == nameof(AngleArc) && pair.Key != "name" && pair.Key != "kind" && pair.Key != "exists" && pair.Key != "visible";
                if (!pixels && !Same(pair.Value, value))
                {
                    differences.Add(where + "." + pair.Key + ": " + Describe(pair.Value) + " here, " + value.GetRawText() + " in the player");
                }
            }
        }
    }

    static bool Same(object expected, JsonElement actual)
    {
        switch (expected)
        {
            case string text:
                // (a label's text has the platform's line breaks here, the browser's in the player)
                return actual.ValueKind == JsonValueKind.String && actual.GetString().Replace("\r\n", "\n") == text.Replace("\r\n", "\n");
            case bool flag:
                return actual.ValueKind == (flag ? JsonValueKind.True : JsonValueKind.False);
            case double number:
                return actual.ValueKind == JsonValueKind.Number && System.Math.Abs(actual.GetDouble() - number) <= Tolerance * System.Math.Max(1, System.Math.Abs(number));
            case object[] pair:
                return actual.ValueKind == JsonValueKind.Array && actual.GetArrayLength() == pair.Length && pair.Select((item, i) => Same(item, actual[i])).All(same => same);
            case List<object[]> list:
                return actual.ValueKind == JsonValueKind.Array && actual.GetArrayLength() == list.Count && list.Select((item, i) => Same(item, actual[i])).All(same => same);
            default:
                return false;
        }
    }

    static string Describe(object value)
    {
        return value switch
        {
            object[] pair => "[" + string.Join(", ", pair.Select(Describe)) + "]",
            List<object[]> list => "[" + string.Join(", ", list.Select(Describe)) + "]",
            double number => number.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            _ => value?.ToString() ?? "null"
        };
    }

    #endregion

    #region The player in headless Edge

    class PlayerDumps
    {
        public List<JsonElement> Load;
        public List<JsonElement> Drag;
        public string Error;
    }

    static Dictionary<string, PlayerDumps> DumpInPlayer(string root, string relativeFolder, List<string> names, bool stop)
    {
        Process server = null;
        if (!IsListening(ServePort))
        {
            server = Process.Start(new ProcessStartInfo("dotnet", "run tools/serve.cs -- Main " + ServePort) { WorkingDirectory = root, UseShellExecute = false });
            WaitFor(ServePort, "the dev server");
        }

        try
        {
            if (!IsListening(EdgePort))
            {
                Process.Start(new ProcessStartInfo("dotnet", "run tools/webauto.cs -- start about:blank 1000 700") { WorkingDirectory = root, UseShellExecute = false });
                WaitFor(EdgePort, "headless Edge");
            }

            Webauto(root, "nav", "http://localhost:" + ServePort + "/Player/dev/index.html?r=parity" + Environment.TickCount);
            Thread.Sleep(2500);
            // (the desktop head has reflection-based JSON serialization off; file names hold no quotes)
            var namesJson = "[" + string.Join(",", names.Select(name => "\"" + name + "\"")) + "]";
            var script = "(async () => { window.parityResult = null; const player = [...LiveGeometry.players.values()][0]; const names = " + namesJson + ";"
                + " const offset = new Point(" + dragOffset.X.ToStringInvariant() + ", " + dragOffset.Y.ToStringInvariant() + "); const result = {};"
                + " for (const name of names) { try { const text = await fetch('/" + relativeFolder + "/' + name).then(r => r.text());"
                + " const drawing = new Drawing(player.canvas); drawing.addFromXml(text); const load = drawing.dump();"
                + " const points = drawing.figures.list.filter(f => f.constructor === FreePoint);"
                + " const dependents = DependencyAlgorithms.findDescendants(f => f.dependents, points); dependents.reverse();"
                + " Actions.move(drawing, points, offset, dependents); result[name] = { load, drag: drawing.dump() }; } catch (e) { result[name] = { error: e.message + ' ' + (e.stack || '').split('\\n')[1] }; } }"
                + " window.parityResult = JSON.stringify(result); })()";
            Webauto(root, "eval", script);
            string json = null;
            for (int i = 0; i < 120 && json == null; i++)
            {
                Thread.Sleep(1000);
                var answer = Webauto(root, "eval", "window.parityResult || ''").Trim();
                if (answer.Length > 0)
                {
                    json = answer;
                }
            }

            if (json == null)
            {
                Console.WriteLine("The player did not finish dumping in two minutes.");
                return null;
            }

            var result = new Dictionary<string, PlayerDumps>();
            using var document = JsonDocument.Parse(json);
            foreach (var property in document.RootElement.EnumerateObject())
            {
                var dumps = new PlayerDumps();
                if (property.Value.TryGetProperty("error", out var error))
                {
                    dumps.Error = error.GetString();
                }
                else
                {
                    dumps.Load = property.Value.GetProperty("load").EnumerateArray().Select(e => e.Clone()).ToList();
                    dumps.Drag = property.Value.GetProperty("drag").EnumerateArray().Select(e => e.Clone()).ToList();
                }

                result[property.Name] = dumps;
            }

            if (stop)
            {
                Webauto(root, "stop");
            }

            return result;
        }
        finally
        {
            if (server != null && !server.HasExited)
            {
                server.Kill(entireProcessTree: true);
            }
        }
    }

    static string Webauto(string root, params string[] arguments)
    {
        var info = new ProcessStartInfo("dotnet") { WorkingDirectory = root, UseShellExecute = false, RedirectStandardOutput = true };
        info.ArgumentList.Add("run");
        info.ArgumentList.Add("tools/webauto.cs");
        info.ArgumentList.Add("--");
        foreach (var argument in arguments)
        {
            info.ArgumentList.Add(argument);
        }

        using var process = Process.Start(info);
        var output = process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        return output;
    }

    static bool IsListening(int port)
    {
        try
        {
            using var client = new TcpClient();
            client.Connect("127.0.0.1", port);
            return true;
        }
        catch (SocketException)
        {
            return false;
        }
    }

    static void WaitFor(int port, string what)
    {
        for (int i = 0; i < 60; i++)
        {
            if (IsListening(port))
            {
                return;
            }

            Thread.Sleep(500);
        }

        throw new InvalidOperationException(what + " did not come up on port " + port + ".");
    }

    #endregion
}
