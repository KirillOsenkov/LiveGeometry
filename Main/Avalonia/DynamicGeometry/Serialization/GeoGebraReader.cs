using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Reads a GeoGebra worksheet (.ggb: a zip with geogebra.xml inside) into a drawing, as far as
/// the figures of this library go. The construction is a list of free elements
/// (&lt;element type="point"&gt; with coords) and commands (&lt;command name="Segment"&gt; with
/// named inputs and outputs, each output followed by its element carrying the style). Every
/// command this reader knows becomes the figure that stands for it here, sometimes with hidden
/// helpers (a circle through three points is the circle around the crossing of two bisectors);
/// a command it doesn't know is reported and its outputs are left out, so what is built on
/// them goes too. Free lines and conics without a command come in as static equations.
/// </summary>
public class GeoGebraReader
{
    /// <summary>The name of the worksheet inside the zip</summary>
    public const string WorksheetEntry = "geogebra.xml";

    /// <summary>Unpacks a .ggb and returns its worksheet; the zip itself has no other use here</summary>
    public static XElement ReadWorksheet(byte[] ggb)
    {
        using var archive = new ZipArchive(new MemoryStream(ggb), ZipArchiveMode.Read);
        var entry = archive.GetEntry(WorksheetEntry);
        if (entry == null)
        {
            throw new InvalidDataException("Not a GeoGebra file: no " + WorksheetEntry + " inside");
        }

        using var stream = entry.Open();
        return XElement.Load(stream);
    }

    public bool IsSuccess
    {
        get
        {
            return messages.Count == 0;
        }
    }

    public string GetErrorReport()
    {
        return "GeoGebra file: " + messages.Count + " thing" + (messages.Count == 1 ? "" : "s") + " left out" + Environment.NewLine
            + string.Join(Environment.NewLine, messages);
    }

    readonly List<string> messages = new List<string>();

    Drawing drawing;

    // GeoGebra label -> figure (the "label" of an element is its name)
    readonly Dictionary<string, IFigure> figures = new Dictionary<string, IFigure>();

    // a polygon's sides are objects of their own in GeoGebra (the angle bisector of two sides);
    // here a side becomes a hidden segment the first time something asks for it
    readonly Dictionary<string, (IPoint, IPoint)> polygonSides = new Dictionary<string, (IPoint, IPoint)>();

    // the view the file was saved in: pixels per unit and where the origin is, for pixel
    // things (a slider on the screen, a text's height)
    double unitLength = Settings.DefaultUnitLength;
    Point origin;
    Size viewSize;
    double defaultFontSize = 16;

    public void ReadDrawing(Drawing drawing, XElement worksheet)
    {
        this.drawing = drawing;
        var font = worksheet.Element("gui")?.Element("font");
        if (font != null && font.ReadDouble("size") > 0)
        {
            defaultFontSize = font.ReadDouble("size");
        }

        ReadView(worksheet.Element("euclidianView"));
        var construction = worksheet.Element("construction");
        if (construction == null)
        {
            return;
        }

        using (drawing.ActionManager.CreateTransaction())
        {
            drawing.ActionManager.RecordingTransaction.IsDelayed = false;
            ReadConstruction(construction);
        }

        new DrawingUpdater().UpdateIfNecessary(drawing);
        drawing.Recalculate();
        ApplyView();
    }

    #region View

    void ReadView(XElement view)
    {
        if (view == null)
        {
            return;
        }

        var coordinates = view.Element("coordSystem");
        if (coordinates != null)
        {
            unitLength = coordinates.ReadDouble("scale");
            if (!unitLength.IsValidPositiveValue())
            {
                unitLength = Settings.DefaultUnitLength;
            }

            origin = new Point(coordinates.ReadDouble("xZero"), coordinates.ReadDouble("yZero"));
        }

        var size = view.Element("size");
        if (size != null)
        {
            viewSize = new Size(size.ReadDouble("width"), size.ReadDouble("height"));
        }

        var settings = view.Element("evSettings");
        if (settings != null)
        {
            bool axes = settings.ReadBool("axes", false);
            bool grid = settings.ReadBool("grid", false);
            drawing.CoordinateGrid.Visible = axes || grid;
            drawing.CoordinateGrid.ShowAxes = axes;
        }

        var background = view.Element("bgColor");
        if (background != null)
        {
            drawing.Background = new SolidColorBrush(ReadColor(background, alpha: 255));
        }
    }

    /// <summary>The same zoom, and the same point in the middle, as the file's window had</summary>
    void ApplyView()
    {
        if (viewSize.Width <= 0 || viewSize.Height <= 0)
        {
            return;
        }

        var center = ToLogical(new Point(viewSize.Width / 2, viewSize.Height / 2));
        drawing.CoordinateSystem.SetView(center, unitLength);
    }

    /// <summary>A pixel of the file's window in the plane</summary>
    Point ToLogical(Point pixel)
    {
        return new Point((pixel.X - origin.X) / unitLength, (origin.Y - pixel.Y) / unitLength);
    }

    #endregion

    #region Construction

    void ReadConstruction(XElement construction)
    {
        // an expression names an element that follows it (a point by coordinates, a text, a
        // function); a command names the elements of its outputs, which follow it too
        XElement pendingExpression = null;
        XElement pendingCommand = null;

        // a text may hang on a point the file defines after it: texts wait for the rest
        var texts = new List<(XElement, XElement)>();
        foreach (var node in construction.Elements())
        {
            try
            {
                switch (node.Name.LocalName)
                {
                    case "expression":
                        pendingExpression = node;
                        break;
                    case "command":
                        pendingCommand = node;
                        pendingExpression = null;
                        RunCommand(node);
                        break;
                    case "element":
                        string label = (string)node.Attribute("label");
                        if (pendingCommand != null && IsOutputOf(pendingCommand, label))
                        {
                            ApplyElement(node);
                            break;
                        }

                        pendingCommand = null;
                        var expression = pendingExpression != null && (string)pendingExpression.Attribute("label") == label ? pendingExpression : null;
                        pendingExpression = null;
                        if ((string)node.Attribute("type") == "text")
                        {
                            texts.Add((node, expression));
                        }
                        else
                        {
                            ReadFreeElement(node, expression);
                        }

                        break;
                }
            }
            catch (Exception ex)
            {
                Report(Describe(node) + ": " + ex.Message);
            }
        }

        foreach (var (node, expression) in texts)
        {
            try
            {
                ReadFreeElement(node, expression);
            }
            catch (Exception ex)
            {
                Report(Describe(node) + ": " + ex.Message);
            }
        }
    }

    static bool IsOutputOf(XElement command, string label)
    {
        var output = command.Element("output");
        return output != null && output.Attributes().Any(a => a.Value == label);
    }

    static string Describe(XElement node)
    {
        return node.Name.LocalName + " " + ((string)node.Attribute("name") ?? (string)node.Attribute("label"));
    }

    void Report(string message)
    {
        messages.Add(message);
    }

    #endregion

    #region Free elements

    /// <summary>An element with no command before it: a free point, a slider, a text, a static line or conic</summary>
    void ReadFreeElement(XElement element, XElement expression)
    {
        string type = (string)element.Attribute("type");
        string label = (string)element.Attribute("label");
        IFigure figure = null;
        switch (type)
        {
            case "point":
                figure = ReadFreePoint(element, expression);
                break;
            case "numeric":
            case "angle":
                figure = ReadNumeric(element);
                break;
            case "text":
                figure = ReadText(element, expression);
                break;
            case "line":
                figure = ReadFreeLine(element);
                break;
            case "conic":
                figure = ReadFreeConic(element);
                break;
            case "function":
                figure = ReadFunction(element, expression);
                break;
            case "boolean":
            case "button":
            case "penstroke":
            case "image":
            case "textfield":
            case "list":
                Report("Skipped " + type + " " + label + ": not supported");
                return;
            default:
                if (expression != null)
                {
                    Report("Skipped " + type + " " + label + " = " + (string)expression.Attribute("exp") + ": not supported");
                }
                else
                {
                    Report("Skipped " + type + " " + label + ": not supported");
                }

                return;
        }

        if (figure != null)
        {
            Register(label, figure);
            ApplyElement(element);
        }
    }

    IFigure ReadFreePoint(XElement element, XElement expression)
    {
        var coordinates = ReadCoordinates(element);
        if (expression != null)
        {
            // (a + x(B), 2), A + (0, 1), t B + (1 - t) A: a point whose coordinates are expressions
            var text = (string)expression.Attribute("exp");
            var point = PointByExpression(text);
            if (point != null)
            {
                return point;
            }

            Report("Point " + (string)element.Attribute("label") + " = " + text + ": expression not understood, a free point instead");
        }

        var free = Factory.CreateFreePoint(drawing, coordinates);
        Actions.Add(drawing, free);
        return free;
    }

    /// <summary>A point by coordinates from a GeoGebra point expression, or null when it doesn't translate</summary>
    PointByCoordinates PointByExpression(string text)
    {
        var vector = TranslateVector(text);
        if (vector == null || !Compiles(vector.Value.Item1) || !Compiles(vector.Value.Item2))
        {
            return null;
        }

        var point = Factory.CreatePointByCoordinates(drawing, vector.Value.Item1, vector.Value.Item2);
        Actions.Add(drawing, point);
        return point;
    }

    /// <summary>
    /// A GeoGebra expression whose value is a point, as our two coordinate expressions: a
    /// literal (a, b), a point by name, sums and differences of those, and a scalar times
    /// one; anything else is null. (A term that is a plain scalar is left to the caller.)
    /// </summary>
    (string, string)? TranslateVector(string text)
    {
        var term = TranslateTerm(text.Trim());
        return term != null && term.Value.Y != null ? (term.Value.X, term.Value.Y) : null;
    }

    /// <summary>A scalar (X only) or a vector (X and Y) of our expression language</summary>
    struct Term
    {
        public string X;
        public string Y;
    }

    Term? TranslateTerm(string text)
    {
        text = text.Trim();
        if (text.Length == 0)
        {
            return null;
        }

        // sums and differences at the top level, minding a leading sign
        var parts = SplitTopLevel(text, '+', '-');
        if (parts.Count > 1)
        {
            Term? sum = null;
            foreach (var (sign, part) in parts)
            {
                var term = TranslateTerm(part);
                if (term == null)
                {
                    return null;
                }

                sum = sum == null ? Signed(term.Value, sign) : Combine(sum.Value, Signed(term.Value, sign));
                if (sum == null)
                {
                    return null;
                }
            }

            return sum;
        }

        var factors = SplitTopLevel(text, '*', '/');
        if (factors.Count > 1)
        {
            Term? product = null;
            foreach (var (sign, part) in factors)
            {
                var factor = TranslateTerm(part);
                if (factor == null)
                {
                    return null;
                }

                if (product == null)
                {
                    product = factor;
                    continue;
                }

                string operation = sign == '/' ? " / " : " * ";
                if (factor.Value.Y == null)
                {
                    product = Scale(product.Value, factor.Value.X, operation);
                }
                else if (product.Value.Y == null && sign != '/')
                {
                    product = Scale(factor.Value, product.Value.X, operation);
                }
                else
                {
                    return null;
                }
            }

            return product;
        }

        if (text.StartsWith("-"))
        {
            var negated = TranslateTerm(text.Substring(1));
            return negated != null ? Signed(negated.Value, '-') : null;
        }

        if (text.StartsWith("(") && text.EndsWith(")"))
        {
            var pair = SplitPoint(text);
            if (pair != null)
            {
                var x = TranslateTerm(pair.Value.Item1);
                var y = TranslateTerm(pair.Value.Item2);
                if (x == null || y == null || x.Value.Y != null || y.Value.Y != null)
                {
                    return null;
                }

                return new Term() { X = x.Value.X, Y = y.Value.X };
            }

            var inner = TranslateTerm(text.Substring(1, text.Length - 2));
            return inner == null ? null : new Term() { X = "(" + inner.Value.X + ")", Y = inner.Value.Y == null ? null : "(" + inner.Value.Y + ")" };
        }

        // a point by name is a vector; anything else is a scalar of our language
        if (IsIdentifier(text.Replace("'", "")) && ResolveArgument(text) is IPoint point && IsIdentifier(point.Name))
        {
            return new Term() { X = point.Name + ".X", Y = point.Name + ".Y" };
        }

        var scalar = TranslateExpression(text);
        return Compiles(scalar) ? new Term() { X = scalar } : null;
    }

    static Term Signed(Term term, char sign)
    {
        if (sign != '-')
        {
            return term;
        }

        return new Term() { X = "-(" + term.X + ")", Y = term.Y == null ? null : "-(" + term.Y + ")" };
    }

    /// <summary>Two vectors or two scalars add; a vector and a scalar don't</summary>
    static Term? Combine(Term a, Term b)
    {
        if ((a.Y == null) != (b.Y == null))
        {
            return null;
        }

        return new Term() { X = a.X + " + " + b.X, Y = a.Y == null ? null : a.Y + " + " + b.Y };
    }

    static Term Scale(Term term, string scalar, string operation)
    {
        return new Term()
        {
            X = "(" + term.X + ")" + operation + "(" + scalar + ")",
            Y = term.Y == null ? null : "(" + term.Y + ")" + operation + "(" + scalar + ")"
        };
    }

    /// <summary>
    /// The operands at the top level of an expression, each with the operator in front of it
    /// (the first with '+', or '-' for a leading minus); one entry when there is no such
    /// operator outside the brackets.
    /// </summary>
    static List<(char, string)> SplitTopLevel(string text, char first, char second)
    {
        var parts = new List<(char, string)>();
        int depth = 0;
        int start = 0;
        char sign = '+';
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '(' || c == '[' || c == '{')
            {
                depth++;
            }
            else if (c == ')' || c == ']' || c == '}')
            {
                depth--;
            }
            else if (depth == 0 && (c == first || c == second) && i > 0 && !IsOperator(text[i - 1]))
            {
                parts.Add((sign, text.Substring(start, i - start)));
                sign = c;
                start = i + 1;
            }
            else if (depth == 0 && i == 0 && c == second)
            {
                // a leading minus: the sign of the first operand
                sign = c;
                start = 1;
            }
        }

        parts.Add((sign, text.Substring(start)));
        if (parts.Count == 1 && sign == '+')
        {
            return parts;
        }

        return parts.Count == 1 && start == 1 ? new List<(char, string)>() { ('+', text) } : parts;
    }

    static bool IsOperator(char c)
    {
        return c == '+' || c == '-' || c == '*' || c == '/' || c == '^' || c == '(' || c == ',';
    }

    /// <summary>"(a, b)" into its two halves, minding the parentheses inside</summary>
    static (string, string)? SplitPoint(string text)
    {
        text = text.Trim();
        if (!text.StartsWith("(") || !text.EndsWith(")"))
        {
            return null;
        }

        int depth = 0;
        for (int i = 1; i < text.Length - 1; i++)
        {
            char c = text[i];
            if (c == '(' || c == '[' || c == '{')
            {
                depth++;
            }
            else if (c == ')' || c == ']' || c == '}')
            {
                depth--;
            }
            else if ((c == ',' || c == '|') && depth == 0)
            {
                return (text.Substring(1, i - 1), text.Substring(i + 1, text.Length - i - 2));
            }
        }

        return null;
    }

    /// <summary>A slider, or a plain number (hidden) that expressions can name</summary>
    IFigure ReadNumeric(XElement element)
    {
        var valueElement = element.Element("value");
        double value = valueElement != null ? valueElement.ReadDouble("val") : 0;
        bool isAngle = (string)element.Attribute("type") == "angle";
        if (isAngle)
        {
            value = value.ToDegrees();
        }

        var sliderElement = element.Element("slider");
        var show = element.Element("show");
        bool shown = show != null && show.ReadBool("object", true);
        if (sliderElement != null && shown)
        {
            var slider = new Slider() { Drawing = drawing };
            slider.Value = value;
            double x = sliderElement.ReadDouble("x");
            double y = sliderElement.ReadDouble("y");
            slider.Position = sliderElement.ReadBool("absoluteScreenLocation", false) ? ToLogical(new Point(x, y)) : new Point(x, y);
            Actions.Add(drawing, slider);
            if (isAngle)
            {
                angleNumbers.Add(slider);
            }

            return slider;
        }

        var number = new Number() { Drawing = drawing, Value = value };
        Actions.Add(drawing, number);
        if (isAngle)
        {
            angleNumbers.Add(number);
        }

        return number;
    }

    // the Numbers that stand for angles: GeoGebra's expressions have them in radians, ours hold degrees
    readonly HashSet<IFigure> angleNumbers = new HashSet<IFigure>();

    IFigure ReadText(XElement element, XElement expression)
    {
        if (expression == null)
        {
            return null;
        }

        // placed by ApplyElement, from the element's startPoint
        return CreateText((string)expression.Attribute("exp"));
    }

    Label CreateText(string expression)
    {
        var label = Factory.CreateLabel(drawing);
        Actions.Add(drawing, label);
        label.Text = TranslateText(expression);
        return label;
    }

    /// <summary>Where the element puts a text: its own coordinates, or a point it hangs on</summary>
    void PlaceText(XElement element, Label label)
    {
        var startPoint = element.Element("startPoint");
        Point place = new Point();
        if (startPoint != null)
        {
            var anchor = (string)startPoint.Attribute("exp");
            if (anchor != null)
            {
                var point = ResolveArgument(anchor) as IPoint;
                if (point != null)
                {
                    place = point.Coordinates;
                }
            }
            else
            {
                place = ReadHomogeneous(startPoint);
            }
        }

        // GeoGebra's point is the start of the first line's baseline; ours is the top-left corner
        double fontSize = ReadFontSize(element);
        label.MoveTo(new Point(place.X, place.Y + fontSize / unitLength));

        // a text with a background color gets a plate (of the paper's color: close enough)
        var background = element.Element("bgColor");
        if (background != null && background.ReadDouble("alpha") > 0)
        {
            label.Backdrop = true;
        }
    }

    double ReadFontSize(XElement element)
    {
        // <font size="16"/> on the element, else the app's font size the gui says
        var font = element.Element("font");
        double size = font != null ? font.ReadDouble("size") : 0;
        return size > 0 ? size : defaultFontSize;
    }

    /// <summary>
    /// A GeoGebra text is an expression over strings: literals in quotes, Name[X] and the
    /// values of objects, joined by +. Here the literals and the names are text and every
    /// value is an expression in brackets that the label evaluates, when one can be written.
    /// </summary>
    string TranslateText(string expression)
    {
        var sb = new StringBuilder();
        foreach (var part in SplitConcatenation(expression))
        {
            var trimmed = part.Trim();
            while (trimmed.StartsWith("(") && trimmed.EndsWith(")") && SplitConcatenation(trimmed.Substring(1, trimmed.Length - 2)).Count == 1)
            {
                trimmed = trimmed.Substring(1, trimmed.Length - 2).Trim();
            }

            if (trimmed.StartsWith("\"") && trimmed.EndsWith("\"") && trimmed.Length >= 2)
            {
                sb.Append(trimmed.Substring(1, trimmed.Length - 2).Replace("\\n", Environment.NewLine));
                continue;
            }

            var name = TryMatchCommand(trimmed, "Name");
            if (name != null)
            {
                var named = ResolveArgument(name);
                sb.Append(named != null ? named.Name : name);
                continue;
            }

            var figure = ResolveArgument(trimmed);
            if (figure != null)
            {
                // Text[t]: another text's words
                if (figure is Label other)
                {
                    sb.Append(other.Text);
                    continue;
                }

                var value = ValueExpressionOf(figure);
                if (value != null)
                {
                    sb.Append("[" + value + "]");
                    continue;
                }

                // no way to say it live (a name with a quote in it): the value as it is now
                var current = CurrentValueOf(figure);
                if (current != null)
                {
                    sb.Append(Math.Round(current.Value, Settings.DisplayDecimals).ToString(CultureInfo.InvariantCulture));
                    continue;
                }
            }

            var translated = TranslateExpression(trimmed);
            if (Compiles(translated))
            {
                sb.Append("[" + translated + "]");
            }
            else
            {
                sb.Append(trimmed);
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// What an expression of ours says for a figure's value, if it has one: a Number by name,
    /// a segment's length, a polygon's area. An angle in degrees for a text (what GeoGebra
    /// prints), in radians inside an expression (what GeoGebra computes with).
    /// </summary>
    string ValueExpressionOf(IFigure figure, bool degrees = true)
    {
        if (figure is INumber && IsIdentifier(figure.Name))
        {
            return angleNumbers.Contains(figure) && !degrees ? "rad(" + figure.Name + ")" : figure.Name;
        }

        if (figure is DistanceMeasurement distance && distance.Dependencies.Count == 2 && distance.Dependencies.All(d => d is IPoint && IsIdentifier(d.Name)))
        {
            return "dist(" + distance.Dependencies[0].Name + ", " + distance.Dependencies[1].Name + ")";
        }

        if (figure is Segment segment && segment.Dependencies.All(d => IsIdentifier(d.Name)))
        {
            return "dist(" + segment.Dependencies[0].Name + ", " + segment.Dependencies[1].Name + ")";
        }

        if (figure is AngleMeasurementBase angle && angle.Dependencies.All(d => IsIdentifier(d.Name)))
        {
            var radians = "oang(" + angle.Dependencies[1].Name + ", " + angle.Dependencies[0].Name + ", " + angle.Dependencies[2].Name + ")";
            return degrees ? "deg(" + radians + ")" : radians;
        }

        if (figure is Polygon polygon && polygon.Dependencies.Count <= 10 && polygon.Dependencies.All(d => d is IPoint && IsIdentifier(d.Name)))
        {
            return "area(" + string.Join(", ", polygon.Dependencies.Select(d => d.Name)) + ")";
        }

        return null;
    }

    static double? CurrentValueOf(IFigure figure)
    {
        switch (figure)
        {
            case INumber number:
                return number.Value;
            case AngleMeasurementBase angle:
                return angle.Measure;
            case AreaMeasurement area:
                return area.Measure;
            case ILengthProvider length:
                return length.Length;
            default:
                return null;
        }
    }

    static bool IsIdentifier(string name)
    {
        return !string.IsNullOrEmpty(name) && name.All(c => char.IsLetterOrDigit(c) || c == '_') && !char.IsDigit(name[0]);
    }

    /// <summary>The top-level operands of a + b + c, quotes and brackets respected</summary>
    static List<string> SplitConcatenation(string text)
    {
        var parts = new List<string>();
        int depth = 0;
        bool inString = false;
        int start = 0;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '"')
            {
                inString = !inString;
            }
            else if (!inString)
            {
                if (c == '(' || c == '[' || c == '{')
                {
                    depth++;
                }
                else if (c == ')' || c == ']' || c == '}')
                {
                    depth--;
                }
                else if (c == '+' && depth == 0)
                {
                    parts.Add(text.Substring(start, i - start));
                    start = i + 1;
                }
            }
        }

        parts.Add(text.Substring(start));
        return parts;
    }

    /// <summary>A line with no construction: a x + b y + c = 0 from its coords, static</summary>
    IFigure ReadFreeLine(XElement element)
    {
        var coordinates = element.Element("coords");
        if (coordinates == null)
        {
            return null;
        }

        var line = Factory.CreateLineByEquation(
            drawing,
            coordinates.ReadDouble("x").ToStringInvariant(),
            coordinates.ReadDouble("y").ToStringInvariant(),
            coordinates.ReadDouble("z").ToStringInvariant());
        Actions.Add(drawing, line);
        return line;
    }

    /// <summary>
    /// A conic with no construction, when it is a circle: the matrix says
    /// A0 x² + A1 y² + A2 + 2 A3 xy + 2 A4 x + 2 A5 y = 0.
    /// </summary>
    IFigure ReadFreeConic(XElement element)
    {
        var matrix = element.Element("matrix");
        if (matrix == null)
        {
            return null;
        }

        double a0 = matrix.ReadDouble("A0");
        double a1 = matrix.ReadDouble("A1");
        double a2 = matrix.ReadDouble("A2");
        double a3 = matrix.ReadDouble("A3");
        double a4 = matrix.ReadDouble("A4");
        double a5 = matrix.ReadDouble("A5");
        if (a0 == 0 || System.Math.Abs(a0 - a1) > 1e-9 * System.Math.Abs(a0) || System.Math.Abs(a3) > 1e-9 * System.Math.Abs(a0))
        {
            Report("Skipped conic " + (string)element.Attribute("label") + ": only circles are supported");
            return null;
        }

        double x = -a4 / a0;
        double y = -a5 / a0;
        double radiusSquared = x * x + y * y - a2 / a0;
        if (radiusSquared <= 0)
        {
            Report("Skipped conic " + (string)element.Attribute("label") + ": not a real circle");
            return null;
        }

        var circle = Factory.CreateCircleByEquation(
            drawing,
            x.ToStringInvariant(),
            y.ToStringInvariant(),
            radiusSquared.SquareRoot().ToStringInvariant());
        Actions.Add(drawing, circle);
        return circle;
    }

    IFigure ReadFunction(XElement element, XElement expression)
    {
        if (expression == null)
        {
            return null;
        }

        var text = (string)expression.Attribute("exp");
        int equals = text.IndexOf('=');
        if (equals >= 0)
        {
            text = text.Substring(equals + 1);
        }

        var function = TranslateExpression(text);
        var result = Compiler.Instance.CompileFunction(drawing, function);
        if (!result.IsSuccess)
        {
            Report("Skipped function " + (string)element.Attribute("label") + " = " + text + ": " + result.GetErrorText());
            return null;
        }

        var graph = new FunctionGraph() { Drawing = drawing };
        graph.FunctionText = function;
        Actions.Add(drawing, graph);
        return graph;
    }

    #endregion

    #region Expressions

    static readonly Regex coordinateFunction = new Regex(@"\b([xy])\(\s*([A-Za-z_][A-Za-z0-9_']*)\s*\)");
    static readonly Regex degrees = new Regex(@"(\d+(?:\.\d+)?)\s*°");
    static readonly Regex implicitMultiplication = new Regex(@"(\d)\s*(?=[A-Za-z(])");

    static readonly Regex identifier = new Regex(@"(?<![\w.'])([A-Za-z_Ͱ-Ͽ][\w']*)(?!\s*[\(\[])");

    /// <summary>
    /// GeoGebra's expression language into ours, as far as they overlap: x(A) is A.X, a
    /// degree sign a factor, the powers written as superscripts as ^, a Distance or Angle
    /// command the function, a figure with a value (a segment, an angle) its value, an angle
    /// in radians as GeoGebra computes with it.
    /// </summary>
    string TranslateExpression(string text)
    {
        text = text.Trim();
        text = coordinateFunction.Replace(text, m => NormalizeName(m.Groups[2].Value) + "." + m.Groups[1].Value.ToUpperInvariant());

        // rad(45), not 45 * pi / 180: with points P and I in the drawing "pi" reads as their distance
        text = degrees.Replace(text, m => "rad(" + m.Groups[1].Value + ")");
        text = text.Replace("²", "^2").Replace("³", "^3").Replace("π", Math.PI.ToString("R", CultureInfo.InvariantCulture)).Replace("ℯ", "e");
        text = ReplaceValueCommands(text);
        text = implicitMultiplication.Replace(text, m => m.Groups[1].Value + " * ");
        text = NormalizeName(text);
        text = identifier.Replace(text, m =>
        {
            // a name of a figure with a value that isn't a Number (a segment, an angle
            // measurement, an angle slider) reads as its value
            var name = m.Groups[1].Value;
            if (name == "pi" || name == "e" || name == "x" || name == "y")
            {
                return name;
            }

            if (figures.TryGetValue(name, out var figure) || polygonSides.ContainsKey(name))
            {
                figure = figure ?? ResolveArgument(name);
                var value = ValueExpressionOf(figure, degrees: false);
                if (value != null && value != name)
                {
                    return value;
                }
            }

            return name;
        });
        return text;
    }

    static readonly Regex valueCommand = new Regex(@"\b(Distance|Length|Angle|Area)\s*\[");

    /// <summary>
    /// Distance[A, B], Length[f], Length[Segment[A, B]], Angle[A, B, C], Area[poly] inside an
    /// expression become what our language says for the value, the brackets matched
    /// properly (an argument may be a command with brackets of its own).
    /// </summary>
    string ReplaceValueCommands(string text)
    {
        for (var match = valueCommand.Match(text); match.Success; match = valueCommand.Match(text))
        {
            int open = match.Index + match.Length - 1;
            int close = MatchingBracket(text, open);
            if (close < 0)
            {
                return text;
            }

            var arguments = SplitArguments(text.Substring(open + 1, close - open - 1)).Select(NormalizeName).ToArray();
            string value = null;
            switch (match.Groups[1].Value)
            {
                case "Distance" when arguments.Length == 2:
                    value = "dist(" + arguments[0] + ", " + arguments[1] + ")";
                    break;
                case "Angle" when arguments.Length == 3:
                    value = "oang(" + arguments[0] + ", " + arguments[1] + ", " + arguments[2] + ")";
                    break;
                case "Length":
                case "Area":
                    if (arguments.Length == 1)
                    {
                        var figure = ResolveArgument(arguments[0]);
                        value = figure != null ? ValueExpressionOf(figure, degrees: false) : null;
                    }

                    break;
            }

            if (value == null)
            {
                return text;
            }

            text = text.Substring(0, match.Index) + "(" + value + ")" + text.Substring(close + 1);
        }

        return text;
    }

    /// <summary>The index of the bracket closing the one at open, or -1</summary>
    static int MatchingBracket(string text, int open)
    {
        int depth = 0;
        for (int i = open; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '[' || c == '(' || c == '{')
            {
                depth++;
            }
            else if (c == ']' || c == ')' || c == '}')
            {
                depth--;
                if (depth == 0)
                {
                    return i;
                }
            }
        }

        return -1;
    }

    /// <summary>
    /// A_{12} is A_12 here: the braces would not parse in an expression, and the name shows
    /// as A₁₂ either way (<see cref="NameDisplay"/>). A_1 stays as it is.
    /// </summary>
    static string NormalizeName(string label)
    {
        return Regex.Replace(label, @"_\{([A-Za-z0-9]+)\}", "_$1");
    }

    bool Compiles(string expression)
    {
        // an empty text compiles to nothing without an error
        return !string.IsNullOrWhiteSpace(expression) && drawing.CompileExpression(expression).IsSuccess;
    }

    #endregion

    #region Commands

    /// <summary>
    /// Runs a command: makes the figure (or figures) that stand for it and registers them under
    /// the output names. A command whose outputs are all unnamed builds nothing (it was there
    /// for something else, a text's anchor, which resolves it again by its text).
    /// </summary>
    void RunCommand(XElement command)
    {
        var outputs = command.Element("output")?.Attributes().Select(a => a.Value).ToArray() ?? Array.Empty<string>();
        if (outputs.All(string.IsNullOrEmpty))
        {
            return;
        }

        string name = (string)command.Attribute("name");
        var inputs = command.Element("input")?.Attributes().Select(a => a.Value).ToArray() ?? Array.Empty<string>();
        var results = CreateCommand(name, inputs, outputs);
        if (results == null)
        {
            Report("Skipped " + name + "[" + string.Join(", ", inputs) + "]: not supported");
            return;
        }

        for (int i = 0; i < outputs.Length && i < results.Count; i++)
        {
            if (!string.IsNullOrEmpty(outputs[i]) && results[i] != null)
            {
                Register(outputs[i], results[i]);
            }
        }
    }

    /// <summary>Command[args] written inline as an argument of another command: built hidden, named after nothing</summary>
    IFigure RunInlineCommand(string name, string[] inputs)
    {
        var results = CreateCommand(name, inputs, Array.Empty<string>());
        var result = results?.FirstOrDefault(r => r != null);
        if (result == null)
        {
            Report("Skipped " + name + "[" + string.Join(", ", inputs) + "]: not supported");
            return null;
        }

        foreach (var figure in results.Where(f => f != null))
        {
            figure.Visible = false;
        }

        return result;
    }

    /// <summary>
    /// The figures for a command, in the order of its outputs (null where an output is not
    /// made); null for a command this reader doesn't know.
    /// </summary>
    List<IFigure> CreateCommand(string name, string[] inputs, string[] outputs)
    {
        switch (name)
        {
            case "Segment":
                return SegmentCommand(inputs);
            case "Line":
                return One(LineCommand(inputs));
            case "Ray":
                return One(Add(Factory.CreateRay(drawing, Points(inputs, count: 2))));
            case "Vector":
                return One(VectorCommand(inputs));
            case "PolyLine":
                return One(Add(Factory.CreatePolyline(drawing, Points(inputs, inputs.Length))));
            case "Polygon":
                return PolygonCommand(inputs, outputs);
            case "Circle":
                return One(CircleCommand(inputs));
            case "Semicircle":
                return One(SemicircleCommand(inputs));
            case "CircularArc":
            case "CircleArc":
                return One(ArcCommand(inputs, Factory.CreateArc));
            case "Centroid":
                return One(CentroidCommand(inputs));
            case "CircularSector":
            case "CircleSector":
                return One(ArcCommand(inputs, Factory.CreateCircleSector));
            case "CircumcircularArc":
            case "CircumcircleArc":
                return One(CircumcircularArcCommand(inputs, Factory.CreateArc));
            case "CircumcircularSector":
            case "CircumcircleSector":
                return One(CircumcircularArcCommand(inputs, Factory.CreateCircleSector));
            case "Ellipse":
                return One(EllipseCommand(inputs));
            case "Midpoint":
            case "Center":
                return One(MidpointCommand(inputs));
            case "Point":
                return One(PointCommand(inputs));
            case "PointIn":
                return One(Add(Factory.CreateFreePoint(drawing, new Point())));
            case "Intersect":
                return IntersectCommand(inputs, outputs);
            case "OrthogonalLine":
            case "PerpendicularLine":
                return One(Add(Factory.CreatePerpendicularLine(drawing, new[] { Line(inputs[1]), PointOf(inputs[0]) })));
            case "LineBisector":
            case "PerpendicularBisector":
                return One(Add(Factory.CreateSegmentBisector(drawing, inputs.Length == 1 ? Ends(inputs[0]) : Points(inputs, count: 2))));
            case "AngularBisector":
            case "AngleBisector":
                return AngleBisectorCommand(inputs);
            case "Tangent":
                return TangentCommand(inputs);
            case "Angle":
                return One(AngleCommand(inputs));
            case "Distance":
            case "Length":
                return One(DistanceCommand(inputs));
            case "Area":
                return One(Add(Factory.CreateAreaMeasurement(drawing, new[] { Resolve(inputs[0]) })));
            case "Text":
                return One(CreateText(inputs[0]));
            case "Mirror":
            case "Reflect":
                return One(Last(Transformer.CreateReflectedFigure(drawing, Resolve(inputs[0]), Resolve(inputs[1]))));
            case "Rotate":
                return One(RotateCommand(inputs));
            case "Translate":
                return One(TranslateCommand(inputs));
            case "Dilate":
                return One(DilateCommand(inputs));
            case "Locus":
                return One(Add(Factory.CreateLocus(drawing, new[] { PointOf(inputs[0]), PointOf(inputs[1]) })));
            default:
                return null;
        }
    }

    static List<IFigure> One(IFigure figure)
    {
        return new List<IFigure>() { figure };
    }

    /// <summary>The figures a transformation made (hidden helpers first) go in; the last is the result</summary>
    IFigure Last(List<IFigure> created)
    {
        foreach (var figure in created)
        {
            Add(figure);
        }

        for (int i = 0; i < created.Count - 1; i++)
        {
            created[i].Visible = false;
        }

        return created[created.Count - 1];
    }

    List<IFigure> SegmentCommand(string[] inputs)
    {
        var start = PointOf(inputs[0]);
        var end = ResolveArgument(inputs[1]) as IPoint;
        if (end != null)
        {
            return One(Add(Factory.CreateSegment(drawing, start, end)));
        }

        // Segment[A, 3]: a segment of a given length, its end free to turn (the second output)
        var length = LengthProvider(inputs[1]);
        var translated = Factory.CreateTranslatedPoint(drawing, start, length, directionSource: null);
        Add(translated);
        var segment = Add(Factory.CreateSegment(drawing, start, translated));
        return new List<IFigure>() { segment, translated };
    }

    IFigure LineCommand(string[] inputs)
    {
        var second = Resolve(inputs[1]);
        if (second is ILine)
        {
            // Line[point, line]: the parallel through the point
            return Add(Factory.CreateParallelLine(drawing, new IFigure[] { second, PointOf(inputs[0]) }));
        }

        if (second is Vector vector)
        {
            var end = Factory.CreateTranslatedPoint(drawing, PointOf(inputs[0]), vector, vector);
            end.Visible = false;
            Add(end);
            return Add(Factory.CreateLineTwoPoints(drawing, new IFigure[] { PointOf(inputs[0]), end }));
        }

        return Add(Factory.CreateLineTwoPoints(drawing, Points(inputs, count: 2)));
    }

    IFigure VectorCommand(string[] inputs)
    {
        if (inputs.Length == 1)
        {
            // Vector[B]: from the origin; the origin is a hidden point by coordinates
            var originPoint = Factory.CreatePointByCoordinates(drawing, "0", "0");
            originPoint.Visible = false;
            Add(originPoint);
            return Add(Factory.CreateVector(drawing, new IFigure[] { originPoint, PointOf(inputs[0]) }));
        }

        return Add(Factory.CreateVector(drawing, Points(inputs, count: 2)));
    }

    /// <summary>
    /// Polygon[A, B, C]: the polygon, then a segment per side in the outputs. Polygon[A, B, n]:
    /// a regular polygon on the side AB, counterclockwise - here a plain polygon whose other
    /// vertices are A turned about the center, so that they keep the names the file gives
    /// them; the center is a point by coordinates, hidden.
    /// </summary>
    List<IFigure> PolygonCommand(string[] inputs, string[] outputs)
    {
        var result = new List<IFigure>();
        List<IPoint> vertices;
        if (inputs.Length == 3 && ResolveArgument(inputs[2]) is not IPoint && double.TryParse(inputs[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double sides) && sides >= 3)
        {
            vertices = RegularPolygonVertices(PointOf(inputs[0]), PointOf(inputs[1]), (int)sides);
        }
        else
        {
            vertices = inputs.Select(PointOf).ToList();
        }

        var polygon = Add(Factory.CreatePolygon(drawing, vertices.Cast<IFigure>().ToList()));
        result.Add(polygon);
        for (int i = 0; i < vertices.Count; i++)
        {
            int outputIndex = 1 + i;
            if (outputIndex < outputs.Length && !string.IsNullOrEmpty(outputs[outputIndex]))
            {
                polygonSides[outputs[outputIndex]] = (vertices[i], vertices[(i + 1) % vertices.Count]);
            }

            result.Add(null);
        }

        // the vertices a regular polygon adds come after its sides
        for (int i = 2; i < vertices.Count; i++)
        {
            result.Add(vertices[i]);
        }

        return result;
    }

    List<IPoint> RegularPolygonVertices(IPoint a, IPoint b, int sides)
    {
        // the center is to the left of A->B, half a side along the perpendicular over tan(pi/n)
        double k = 0.5 / System.Math.Tan(Math.PI / sides);
        string kText = k.ToStringInvariant();
        var center = Factory.CreatePointByCoordinates(
            drawing,
            "(" + a.Name + ".X + " + b.Name + ".X) / 2 - (" + b.Name + ".Y - " + a.Name + ".Y) * " + kText,
            "(" + a.Name + ".Y + " + b.Name + ".Y) / 2 + (" + b.Name + ".X - " + a.Name + ".X) * " + kText);
        center.Visible = false;
        Add(center);
        var vertices = new List<IPoint>() { a, b };
        for (int i = 2; i < sides; i++)
        {
            var vertex = Factory.CreateRotatedPoint(drawing, new IFigure[] { a, center }, i * 360.0 / sides);
            Add(vertex);
            vertices.Add(vertex);
        }

        return vertices;
    }

    IFigure CircleCommand(string[] inputs)
    {
        if (inputs.Length == 3)
        {
            // through three points: around the crossing of two perpendicular bisectors
            var points = Points(inputs, count: 3);
            var center = Circumcenter(points);
            return Add(Factory.CreateCircle(drawing, new[] { center, points[0] }));
        }

        var second = ResolveArgument(inputs[1]);
        if (second is IPoint)
        {
            return Add(Factory.CreateCircle(drawing, Points(inputs, count: 2)));
        }

        // a segment, a number, a slider: the radius, then the center
        return Add(Factory.CreateCircleByRadius(drawing, new[] { LengthProvider(inputs[1]), PointOf(inputs[0]) }));
    }

    IPoint Circumcenter(IList<IFigure> points)
    {
        var bisector1 = Factory.CreateSegmentBisector(drawing, new[] { points[0], points[1] });
        var bisector2 = Factory.CreateSegmentBisector(drawing, new[] { points[1], points[2] });
        bisector1.Visible = false;
        bisector2.Visible = false;
        Add(bisector1);
        Add(bisector2);
        var center = Factory.CreateIntersectionPoint(drawing, bisector1, bisector2, new Point());
        center.Visible = false;
        Add(center);
        return center;
    }

    IFigure SemicircleCommand(string[] inputs)
    {
        var points = Points(inputs, count: 2);
        var center = Factory.CreateMidPoint(drawing, points);
        center.Visible = false;
        Add(center);
        return Add(Factory.CreateArc(drawing, new[] { center, points[0], points[1] }));
    }

    /// <summary>CircularArc[center, A, B]: counterclockwise from A towards B, which is the same here</summary>
    IFigure ArcCommand(string[] inputs, Func<Drawing, IList<IFigure>, IFigure> create)
    {
        return Add(create(drawing, Points(inputs, count: 3)));
    }

    /// <summary>The arc through three points, from the first through the second to the third</summary>
    IFigure CircumcircularArcCommand(string[] inputs, Func<Drawing, IList<IFigure>, IFigure> create)
    {
        var points = Points(inputs, count: 3);
        var center = Circumcenter(points);
        var arc = (IArc)create(drawing, new[] { center, points[0], points[2] });
        var middle = ((IPoint)points[1]).Coordinates;
        arc.Clockwise = Math.OAngle(((IPoint)points[0]).Coordinates, center.Coordinates, middle)
            > Math.OAngle(((IPoint)points[0]).Coordinates, center.Coordinates, ((IPoint)points[2]).Coordinates);
        return Add(arc);
    }

    /// <summary>
    /// Ellipse[F1, F2, P]: foci and a point on it. Ours is center, end of the long axis, end
    /// of the short axis: a midpoint and two points by coordinates, hidden.
    /// </summary>
    IFigure EllipseCommand(string[] inputs)
    {
        var f1 = PointOf(inputs[0]);
        var f2 = PointOf(inputs[1]);
        var third = ResolveArgument(inputs[2]);
        string a;
        if (third is IPoint p)
        {
            a = "(dist(" + p.Name + ", " + f1.Name + ") + dist(" + p.Name + ", " + f2.Name + ")) / 2";
        }
        else
        {
            var length = LengthProvider(inputs[2]);
            a = length.Name;
        }

        foreach (var named in new IFigure[] { f1, f2 })
        {
            if (!IsIdentifier(named.Name))
            {
                Report("Ellipse over " + named.Name + ": the name can't be said in an expression");
                return null;
            }
        }

        string c = "dist(" + f1.Name + ", " + f2.Name + ") / 2";
        string ux = "(" + f2.Name + ".X - " + f1.Name + ".X) / dist(" + f1.Name + ", " + f2.Name + ")";
        string uy = "(" + f2.Name + ".Y - " + f1.Name + ".Y) / dist(" + f1.Name + ", " + f2.Name + ")";
        string cx = "(" + f1.Name + ".X + " + f2.Name + ".X) / 2";
        string cy = "(" + f1.Name + ".Y + " + f2.Name + ".Y) / 2";
        string b = "sqrt((" + a + ")^2 - (" + c + ")^2)";
        var center = Factory.CreateMidPoint(drawing, new IFigure[] { f1, f2 });
        var major = Factory.CreatePointByCoordinates(drawing, cx + " + (" + a + ") * " + ux, cy + " + (" + a + ") * " + uy);
        var minor = Factory.CreatePointByCoordinates(drawing, cx + " - (" + b + ") * " + uy, cy + " + (" + b + ") * " + ux);
        center.Visible = false;
        major.Visible = false;
        minor.Visible = false;
        Add(center);
        Add(major);
        Add(minor);
        return Add(Factory.CreateEllipse(drawing, new IFigure[] { center, major, minor }));
    }

    IFigure MidpointCommand(string[] inputs)
    {
        if (inputs.Length == 1)
        {
            var figure = Resolve(inputs[0]);
            if (figure is ICircle circle)
            {
                // Center[c]: the center of a circle by two points is its first point
                if (circle.Dependencies.Count > 0 && circle.Dependencies[circle is CircleByRadius ? circle.Dependencies.Count - 1 : 0] is IPoint centerPoint)
                {
                    return centerPoint;
                }

                return null;
            }

            return Add(Factory.CreateMidPoint(drawing, Ends(inputs[0])));
        }

        return Add(Factory.CreateMidPoint(drawing, Points(inputs, count: 2)));
    }

    /// <summary>Point[figure]: on the figure, at the saved coordinates once the element is read</summary>
    IFigure PointCommand(string[] inputs)
    {
        var figure = Resolve(inputs[0]);
        if (!(figure is ILinearFigure))
        {
            Report("Point on " + inputs[0] + ": a point can't be on a " + figure.GetType().Name);
            return null;
        }

        return Add(Factory.CreatePointOnFigure(drawing, figure, parameter: 0));
    }

    /// <summary>
    /// Intersect[a, b]: one output per crossing, each picked by the saved coordinates of its
    /// element - which come after the command, so the choice is made when the element is read
    /// (<see cref="ApplyElement"/> moves an intersection point to its coordinates).
    /// </summary>
    List<IFigure> IntersectCommand(string[] inputs, string[] outputs)
    {
        var figure1 = Resolve(inputs[0]);
        var figure2 = Resolve(inputs[1]);
        var algorithms = IntersectionPoint.GetAlgorithms(figure1, figure2);
        if (algorithms.Length == 0)
        {
            Report("Intersect[" + inputs[0] + ", " + inputs[1] + "]: those two can't be intersected here");
            return null;
        }

        // by index until the elements say where they are (a crossing that doesn't exist
        // right now has no coordinates in the file)
        var result = new List<IFigure>();
        int count = System.Math.Max(1, outputs.Length);
        for (int i = 0; i < count; i++)
        {
            var point = new IntersectionPoint(new Point(), new List<IFigure> { figure1, figure2 }) { Drawing = drawing };
            point.SetAlgorithm(algorithms[System.Math.Min(i, algorithms.Length - 1)]);
            result.Add(Add(point));
        }

        return result;
    }

    /// <summary>
    /// AngularBisector[A, B, C]: the bisector at B, a whole line. AngularBisector[g, h]: the two
    /// bisectors of the lines: the interior one of the angle at their crossing, and the one
    /// perpendicular to it.
    /// </summary>
    List<IFigure> AngleBisectorCommand(string[] inputs)
    {
        if (inputs.Length == 3)
        {
            var bisector = Factory.CreateAngleBisector(drawing, new[] { PointOf(inputs[1]), PointOf(inputs[0]), PointOf(inputs[2]) });
            bisector.Interior = true;
            bisector.IsLine = true;
            return One(Add(bisector));
        }

        var line1 = Line(inputs[0]);
        var line2 = Line(inputs[1]);
        var vertex = SharedPoint(line1, line2);
        IFigure side1;
        IFigure side2;
        if (vertex != null)
        {
            side1 = OtherPoint(line1, vertex);
            side2 = OtherPoint(line2, vertex);
        }
        else
        {
            var crossing = Factory.CreateIntersectionPoint(drawing, line1, line2, new Point());
            crossing.Visible = false;
            Add(crossing);
            vertex = crossing;
            side1 = line1.Dependencies.OfType<IPoint>().FirstOrDefault();
            side2 = line2.Dependencies.OfType<IPoint>().FirstOrDefault();
            if (side1 == null || side2 == null)
            {
                Report("AngularBisector[" + inputs[0] + ", " + inputs[1] + "]: lines without points on them");
                return null;
            }
        }

        var first = Factory.CreateAngleBisector(drawing, new[] { vertex, side1, side2 });
        first.Interior = true;
        first.IsLine = true;
        Add(first);
        var second = Add(Factory.CreatePerpendicularLine(drawing, new IFigure[] { first, vertex }));
        return new List<IFigure>() { first, second };
    }

    static IPoint SharedPoint(IFigure line1, IFigure line2)
    {
        return line1.Dependencies.OfType<IPoint>().FirstOrDefault(p => line2.Dependencies.Contains(p));
    }

    static IPoint OtherPoint(IFigure line, IPoint point)
    {
        return line.Dependencies.OfType<IPoint>().FirstOrDefault(p => p != point) ?? point;
    }

    /// <summary>
    /// Tangent[P, c]: the two tangents from a point to a circle, through the points where the
    /// circle on the diameter from P to the center crosses it (hidden).
    /// </summary>
    List<IFigure> TangentCommand(string[] inputs)
    {
        var point = PointOf(inputs[0]);
        var circle = Resolve(inputs[1]) as ICircle;
        if (circle == null)
        {
            Report("Tangent[" + inputs[0] + ", " + inputs[1] + "]: only tangents to a circle are supported");
            return null;
        }

        IFigure center = circle is CircleByRadius
            ? circle.Dependencies[circle.Dependencies.Count - 1]
            : circle.Dependencies.OfType<IPoint>().FirstOrDefault();
        if (center == null)
        {
            Report("Tangent[" + inputs[0] + ", " + inputs[1] + "]: the circle has no center point");
            return null;
        }

        var middle = Factory.CreateMidPoint(drawing, new IFigure[] { point, center });
        middle.Visible = false;
        Add(middle);
        var thales = Factory.CreateCircle(drawing, new IFigure[] { middle, point });
        thales.Visible = false;
        Add(thales);
        var result = new List<IFigure>();
        var algorithms = IntersectionPoint.GetAlgorithms(circle, thales);
        for (int i = 0; i < 2; i++)
        {
            var touch = new IntersectionPoint(new Point(), new List<IFigure> { (IFigure)circle, thales }) { Drawing = drawing };
            touch.SetAlgorithm(algorithms[System.Math.Min(i, algorithms.Length - 1)]);
            touch.Visible = false;
            Add(touch);
            result.Add(Add(Factory.CreateLineTwoPoints(drawing, new IFigure[] { point, touch })));
        }

        return result;
    }

    /// <summary>
    /// Angle[A, B, C]: at B, from A to C counterclockwise; the mark comes along as an
    /// <see cref="AngleArc"/>. Which of the two angles at B the file means is fixed by the
    /// element's angleStyle, read after (<see cref="ApplyElement"/>).
    /// </summary>
    IFigure AngleCommand(string[] inputs)
    {
        if (inputs.Length != 3)
        {
            return null;
        }

        var angle = Factory.CreateAngleMeasurement(drawing, new[] { PointOf(inputs[1]), PointOf(inputs[0]), PointOf(inputs[2]) });
        Add(angle);
        var arc = Factory.CreateAngleArc(drawing, new[] { PointOf(inputs[1]), PointOf(inputs[0]), PointOf(inputs[2]) });
        arc.ArcCount = 1;
        Add(arc);
        return angle;
    }

    IFigure DistanceCommand(string[] inputs)
    {
        if (inputs.Length == 1)
        {
            return Add(Factory.CreateDistanceMeasurement(drawing, new[] { Resolve(inputs[0]) }));
        }

        return Add(Factory.CreateDistanceMeasurement(drawing, Points(inputs, count: 2)));
    }

    /// <summary>Rotate[figure, angle, center]; without a center, about the origin</summary>
    IFigure RotateCommand(string[] inputs)
    {
        var source = Resolve(inputs[0]);
        IFigure center = inputs.Length > 2 ? PointOf(inputs[2]) : OriginPoint();
        var angle = ResolveArgument(inputs[1]);
        if (!(angle is IAngleProvider))
        {
            angle = NumberArgument(inputs[1], isAngle: true);
        }

        return Last(Transformer.CreateRotatedFigure(drawing, source, center, angle, angle: 0));
    }

    IFigure TranslateCommand(string[] inputs)
    {
        var source = Resolve(inputs[0]);
        var vector = Resolve(inputs[1]) as Vector;
        if (vector == null)
        {
            Report("Translate[" + inputs[0] + ", " + inputs[1] + "]: only by a vector");
            return null;
        }

        return Last(Transformer.CreateTranslatedFigure(drawing, source, vector, vector));
    }

    IFigure DilateCommand(string[] inputs)
    {
        var source = Resolve(inputs[0]);
        IFigure center = inputs.Length > 2 ? PointOf(inputs[2]) : OriginPoint();
        var factor = LengthProvider(inputs[1]);
        return Last(Transformer.CreateDilatedFigure(drawing, source, center, factor, lengthProvider2: null, factor: 0));
    }

    /// <summary>Centroid[polygon]: the mean of the vertices, a point by coordinates</summary>
    IFigure CentroidCommand(string[] inputs)
    {
        var polygon = Resolve(inputs[0]);
        var vertices = polygon.Dependencies.OfType<IPoint>().ToList();
        if (vertices.Count == 0 || vertices.Any(v => !IsIdentifier(v.Name)))
        {
            Report("Centroid[" + inputs[0] + "]: no vertices to average");
            return null;
        }

        string count = vertices.Count.ToString(CultureInfo.InvariantCulture);
        return Add(Factory.CreatePointByCoordinates(
            drawing,
            "(" + string.Join(" + ", vertices.Select(v => v.Name + ".X")) + ") / " + count,
            "(" + string.Join(" + ", vertices.Select(v => v.Name + ".Y")) + ") / " + count));
    }

    IPoint OriginPoint()
    {
        var point = Factory.CreatePointByCoordinates(drawing, "0", "0");
        point.Visible = false;
        Add(point);
        return point;
    }

    #endregion

    #region Arguments

    IFigure Add(IFigure figure)
    {
        Actions.Add(drawing, figure);
        return figure;
    }

    /// <summary>The figure takes the file's name as it is: A_1 shows as A₁ here too (<see cref="NameDisplay"/>)</summary>
    void Register(string label, IFigure figure)
    {
        figures[label] = figure;
        var name = NormalizeName(label);
        if (figure.Name != name)
        {
            figure.Name = name;
        }
    }

    /// <summary>A figure by its GeoGebra name, or an inline command, or nothing for a number</summary>
    IFigure ResolveArgument(string argument)
    {
        argument = argument.Trim();
        if (figures.TryGetValue(argument, out var figure))
        {
            return figure;
        }

        if (polygonSides.TryGetValue(argument, out var side))
        {
            var segment = Factory.CreateSegment(drawing, side.Item1, side.Item2);
            segment.Visible = false;
            Add(segment);
            figures[argument] = segment;
            return segment;
        }

        var inline = ParseCommand(argument);
        if (inline != null)
        {
            var created = RunInlineCommand(inline.Value.Item1, inline.Value.Item2);
            if (created != null)
            {
                figures[argument] = created;
            }

            return created;
        }

        // the axes are figures in GeoGebra: hidden lines by equation here
        if (argument == "xAxis" || argument == "yAxis")
        {
            var axis = argument == "xAxis"
                ? Factory.CreateLineByEquation(drawing, "0", "1", "0")
                : Factory.CreateLineByEquation(drawing, "1", "0", "0");
            axis.Visible = false;
            Add(axis);
            figures[argument] = axis;
            return axis;
        }

        // (a, b): a point by coordinates, hidden
        if (SplitPoint(argument) != null)
        {
            var point = PointByExpression(argument);
            if (point != null)
            {
                point.Visible = false;
                figures[argument] = point;
            }

            return point;
        }

        return null;
    }

    IFigure Resolve(string argument)
    {
        var figure = ResolveArgument(argument);
        if (figure == null)
        {
            throw new InvalidDataException("'" + argument + "' is not a figure that was read");
        }

        return figure;
    }

    IPoint PointOf(string argument)
    {
        var figure = Resolve(argument);
        if (figure is IPoint point)
        {
            return point;
        }

        throw new InvalidDataException("'" + argument + "' is not a point");
    }

    IFigure Line(string argument)
    {
        var figure = Resolve(argument);
        if (figure is ILine)
        {
            return figure;
        }

        throw new InvalidDataException("'" + argument + "' is not a line");
    }

    IList<IFigure> Points(string[] inputs, int count)
    {
        var result = new List<IFigure>();
        for (int i = 0; i < count; i++)
        {
            result.Add(PointOf(inputs[i]));
        }

        return result;
    }

    /// <summary>The two points of a segment</summary>
    IList<IFigure> Ends(string argument)
    {
        var figure = Resolve(argument);
        var ends = figure.Dependencies.OfType<IPoint>().Take(2).Cast<IFigure>().ToList();
        if (ends.Count != 2)
        {
            throw new InvalidDataException("'" + argument + "' has no two points");
        }

        return ends;
    }

    /// <summary>Something with a length: a segment, a slider, a number; a typed value becomes a Number</summary>
    IFigure LengthProvider(string argument)
    {
        var figure = ResolveArgument(argument);
        if (figure is ILengthProvider)
        {
            return figure;
        }

        if (figure != null)
        {
            throw new InvalidDataException("'" + argument + "' has no length");
        }

        return NumberArgument(argument, isAngle: false);
    }

    /// <summary>
    /// A number typed into a command - 3, 45°, 0.2 * a, -α - as a figure with that value: a
    /// Number when it is constant, a hidden label evaluating the expression when it depends
    /// on figures (a label is a length and an angle provider). Angles are radians in
    /// GeoGebra's expressions; a Number holds degrees, a label's value is what it says.
    /// </summary>
    IFigure NumberArgument(string text, bool isAngle)
    {
        text = text.Trim();
        if (text == "°")
        {
            // Rotate[A, °, B]: an angle nobody typed
            text = "0";
        }

        var expression = TranslateExpression(text);
        var compiled = string.IsNullOrWhiteSpace(expression) ? null : drawing.CompileExpression(expression);
        if (compiled == null || !compiled.IsSuccess)
        {
            throw new InvalidDataException("'" + text + "' is not a number");
        }

        if (compiled.Dependencies.Count == 0)
        {
            double value = compiled.Expression();
            return Add(Number.CreateAuxiliary(drawing, isAngle ? value.ToDegrees() : value));
        }

        var label = Factory.CreateLabel(drawing);
        label.Visible = false;
        label.Auxiliary = true;
        Add(label);
        label.DecimalsToShow = 10;
        label.Text = "[" + expression + "]";
        return label;
    }

    /// <summary>Command[a, b] or Command(a, b) into its name and arguments</summary>
    static (string, string[])? ParseCommand(string text)
    {
        var match = Regex.Match(text, @"^([A-Za-z]+)\s*[\[(](.*)[\])]$");
        if (!match.Success)
        {
            return null;
        }

        return (match.Groups[1].Value, SplitArguments(match.Groups[2].Value));
    }

    static string TryMatchCommand(string text, string name)
    {
        var command = ParseCommand(text);
        return command != null && command.Value.Item1 == name && command.Value.Item2.Length == 1 ? command.Value.Item2[0] : null;
    }

    static string[] SplitArguments(string text)
    {
        var parts = new List<string>();
        int depth = 0;
        int start = 0;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '(' || c == '[' || c == '{')
            {
                depth++;
            }
            else if (c == ')' || c == ']' || c == '}')
            {
                depth--;
            }
            else if (c == ',' && depth == 0)
            {
                parts.Add(text.Substring(start, i - start).Trim());
                start = i + 1;
            }
        }

        parts.Add(text.Substring(start).Trim());
        return parts.ToArray();
    }

    #endregion

    #region Elements

    /// <summary>
    /// What the element after a command says about its figure: whether it shows, its label,
    /// color, size, and for a point its coordinates (which place a point on a figure and pick
    /// the crossing an intersection point is).
    /// </summary>
    void ApplyElement(XElement element)
    {
        string label = (string)element.Attribute("label");
        if (!figures.TryGetValue(label, out var figure))
        {
            return;
        }

        // a number (a distance, an angle's value) is in the view only when the file says so
        var show = element.Element("show");
        string type = (string)element.Attribute("type");
        bool visible = show != null ? show.ReadBool("object", true) : type != "numeric" && type != "angle";
        bool showLabel = show != null && show.ReadBool("label", false);
        if (!(figure is INumber) || figure is Slider)
        {
            // first: a point's label is placed by its size (a plain number has no shape to style)
            ApplyStyle(element, figure);
        }

        if (figure is PointBase point)
        {
            ApplyPointElement(element, point);
            // labelMode: 0 name, 1 name and value, 2 value, 3 caption, 9 caption and value
            var labelMode = element.Element("labelMode");
            int mode = labelMode != null ? (int)labelMode.ReadDouble("val") : 0;
            point.Visible = visible;
            point.ShowName = visible && showLabel && mode != 2;
            point.ShowCoordinates = visible && showLabel && (mode == 1 || mode == 2 || mode == 9);
            if (point.Label != null)
            {
                // GeoGebra writes the name in the point's color
                var color = ReadObjectColor(element);
                if (color != null)
                {
                    point.Label.Style = TextStyleFor(color.Value, defaultFontSize);
                }

                // GeoGebra starts the name's baseline a point's radius to the upper right of
                // the point, plus labelOffset; ours is the top-left corner of the text
                var labelOffset = element.Element("labelOffset");
                double offsetX = labelOffset != null ? labelOffset.ReadDouble("x") : 0;
                double offsetY = labelOffset != null ? labelOffset.ReadDouble("y") : 0;
                double radius = point.Style is PointStyle pointStyle ? pointStyle.Size / 2 : 5;
                point.Label.Offset = new Point(radius + offsetX, -radius + offsetY - defaultFontSize);
            }
        }
        else if (figure is AngleMeasurement angle)
        {
            ApplyAngleElement(element, angle);
            angle.Visible = visible;
        }
        else if (figure is Label text)
        {
            PlaceText(element, text);
            text.Visible = visible;
        }
        else
        {
            figure.Visible = visible;
        }
    }

    void ApplyPointElement(XElement element, PointBase point)
    {
        var coordinates = element.Element("coords");
        if (coordinates == null)
        {
            return;
        }

        var saved = ReadHomogeneous(coordinates);
        if (!saved.Exists())
        {
            return;
        }

        // only points that dragging moves on their own: moving a dependent point would move
        // what it is built on
        if (point is IntersectionPoint intersection)
        {
            intersection.PickNearest(saved);
        }
        else if ((point is FreePoint && !(point is PointByCoordinates)) || (point is TranslatedPoint translated && translated.HasFreedom))
        {
            point.MoveTo(saved);
        }
    }

    /// <summary>
    /// angleStyle: 0 counterclockwise as given (may be reflex), 1 never reflex, 2 always
    /// reflex, 3 unbounded. Ours goes counterclockwise from its second point to its third:
    /// swapping them gives the other angle.
    /// </summary>
    void ApplyAngleElement(XElement element, AngleMeasurement angle)
    {
        var style = element.Element("angleStyle");
        int angleStyle = style != null ? (int)style.ReadDouble("val") : 0;
        bool reflex = Math.OAngle(angle.Point(1), angle.Point(0), angle.Point(2)) > Math.PI;
        if ((angleStyle == 1 && reflex) || (angleStyle == 2 && !reflex))
        {
            AngleArc.ConvertToOpposite(angle);
        }

        var arc = AngleArc.FindCompanion(angle) as AngleArc;
        if (arc != null)
        {
            var size = element.Element("arcSize");
            if (size != null && size.ReadDouble("val") >= 10)
            {
                arc.Size = System.Math.Min(size.ReadDouble("val"), 100);
            }

            var show = element.Element("show");
            arc.Visible = show == null || show.ReadBool("object", true);
            var color = ReadObjectColor(element);
            if (color != null)
            {
                var (stroke, width, dash) = ReadStroke(element, color.Value);
                arc.Style = drawing.StyleManager.FindExistingOrAddNew(new ShapeStyle()
                {
                    Color = stroke,
                    StrokeWidth = width,
                    Dash = dash,
                    IsFilled = false
                });
            }

            var decoration = element.Element("decoration");
            if (decoration != null)
            {
                // 0 none, 1 two arcs, 2 three arcs, 3-5 ticks, 6-7 arrows
                int type = (int)decoration.ReadDouble("type");
                arc.ArcCount = type == 1 ? 2 : type == 2 ? 3 : 1;
            }
        }
    }

    static Point ReadCoordinates(XElement element)
    {
        var coordinates = element.Element("coords");
        return coordinates != null ? ReadHomogeneous(coordinates) : new Point();
    }

    /// <summary>x, y, z homogeneous: the point is (x/z, y/z); z is 1 for most</summary>
    static Point ReadHomogeneous(XElement coordinates)
    {
        double x = coordinates.ReadDouble("x");
        double y = coordinates.ReadDouble("y");
        double z = coordinates.Attribute("z") != null ? coordinates.ReadDouble("z") : 1;
        if (z == 0)
        {
            return new Point(double.NaN, double.NaN);
        }

        return new Point(x / z, y / z);
    }

    #endregion

    #region Styles

    // The drawing keeps GeoGebra's look, since it was made in it: every element carries its
    // color, size, thickness and opacities, and their defaults (blue free points, gray
    // dependent ones and lines, a polygon filled a tenth) are what the file says too.

    /// <summary>The stroke of a line or an outline: the object's color at the line's opacity, thickness/2 pixels wide, dashed by type</summary>
    static (Color, double, LineDash) ReadStroke(XElement element, Color color)
    {
        var lineStyle = element.Element("lineStyle");
        double thickness = lineStyle != null ? lineStyle.ReadDouble("thickness") : 5;
        double width = System.Math.Max(thickness / 2, 0.5);
        int type = lineStyle != null ? (int)lineStyle.ReadDouble("type") : 0;
        var dash = type == 10 || type == 15 ? LineDash.Dash : type == 20 ? LineDash.Dot : type == 30 ? LineDash.DashDot : LineDash.Solid;
        double opacity = lineStyle != null && lineStyle.Attribute("opacity") != null ? lineStyle.ReadDouble("opacity") / 255 : 1;
        var stroke = Color.FromArgb((byte)(255 * opacity), color.R, color.G, color.B);
        return (stroke, width, dash);
    }

    /// <summary>
    /// The fill: the object's color at objColor's alpha (0 for lines and points, 0.1 for a
    /// polygon). A pattern fill (fillType 1 hatch, 2 crosshatch, 3 chessboard, 4 dots...) has
    /// no alpha and no counterpart here: a translucent fill of the color stands in for it.
    /// </summary>
    static (Color, bool) ReadFill(XElement element, Color color)
    {
        var objectColor = element.Element("objColor");
        double alpha = objectColor != null ? objectColor.ReadDouble("alpha") : 0;
        if (objectColor != null && objectColor.ReadDouble("fillType") > 0 && alpha == 0)
        {
            alpha = 0.4;
        }

        return (Color.FromArgb((byte)(255 * System.Math.Min(1, alpha)), color.R, color.G, color.B), alpha > 0);
    }

    IFigureStyle TextStyleFor(Color color, double fontSize)
    {
        return drawing.StyleManager.FindExistingOrAddNew(new TextStyle()
        {
            Color = color,
            FontSize = fontSize,
            FontFamily = new FontFamily("Arial")
        });
    }

    void ApplyStyle(XElement element, IFigure figure)
    {
        var color = ReadObjectColor(element);
        if (color == null)
        {
            return;
        }

        var (stroke, width, dash) = ReadStroke(element, color.Value);
        var (fill, filled) = ReadFill(element, color.Value);
        IFigureStyle style;
        if (figure is IPoint)
        {
            style = ReadPointStyle(element, color.Value);
        }
        else if (figure is LabelBase)
        {
            style = TextStyleFor(color.Value, ReadFontSize(element));
        }
        else if (figure is IShapeWithInterior || figure is Bezier)
        {
            style = new ShapeStyle()
            {
                Fill = new SolidColorBrush(fill),
                IsFilled = filled,
                Color = stroke,
                StrokeWidth = width,
                Dash = dash
            };
        }
        else if (figure is ILinearFigure || figure is Vector || figure is Slider)
        {
            style = new LineStyle() { Color = stroke, StrokeWidth = width, Dash = dash };
            if (figure is Slider slider)
            {
                // GeoGebra's slider is a gray bar with a dark knob
                var knob = drawing.StyleManager.FindExistingOrAddNew(new PointStyle()
                {
                    Fill = new SolidColorBrush(color.Value),
                    Color = Darker(color.Value),
                    StrokeWidth = 1,
                    Size = 10
                });
                slider.Knob.Style = knob;
                slider.Anchor.Style = knob;
            }
        }
        else
        {
            return;
        }

        figure.Style = drawing.StyleManager.FindExistingOrAddNew(style);
    }

    /// <summary>
    /// A GeoGebra point is 2 * pointSize across in its color, with a rim a shade darker.
    /// pointStyle: 0 dot, 1 cross, 2 ring, 3 plus, 4 diamond, 5 hollow diamond, 6-9 triangles
    /// pointing up, down, right, left, 10 a dot without the rim. A cross and a plus are
    /// drawn as the characters, the ring and the hollow diamond as unfilled shapes.
    /// </summary>
    static PointStyle ReadPointStyle(XElement element, Color color)
    {
        var sizeElement = element.Element("pointSize");
        double size = sizeElement != null ? sizeElement.ReadDouble("val") : 5;
        var shapeElement = element.Element("pointStyle");
        int shape = shapeElement != null ? (int)shapeElement.ReadDouble("val") : 0;
        var style = new PointStyle()
        {
            Fill = new SolidColorBrush(color),
            Color = Darker(color),
            StrokeWidth = 1,
            Shape = shape == 4 || shape == 5 ? PointShape.Diamond : shape >= 6 && shape <= 9 ? PointShape.Triangle : PointShape.Circle
        };
        switch (shape)
        {
            case 1:
                style.Character = "×";
                break;
            case 3:
                style.Character = "+";
                break;
            case 2:
            case 5:
                style.IsFilled = false;
                style.Color = color;
                style.StrokeWidth = System.Math.Max(1.5, size / 2.5);
                break;
            case 10:
                style.StrokeWidth = 0;
                break;
        }

        // the size after the character: a style keeps one size for the shape and one for the character
        style.Size = System.Math.Max(3, size * 2) + (style.Character != null ? 6 : 0);
        return style;
    }

    static Color Darker(Color color)
    {
        return Color.FromRgb((byte)(color.R * 0.7), (byte)(color.G * 0.7), (byte)(color.B * 0.7));
    }

    /// <summary>The object's color; a dynamic color (fractions in dynamicr/g/b, what the object shows right now) wins over the static one</summary>
    static Color? ReadObjectColor(XElement element)
    {
        var color = element.Element("objColor");
        if (color == null)
        {
            return null;
        }

        if (color.Attribute("dynamicr") != null)
        {
            return Color.FromRgb(
                (byte)(255 * color.ReadDouble("dynamicr")),
                (byte)(255 * color.ReadDouble("dynamicg")),
                (byte)(255 * color.ReadDouble("dynamicb")));
        }

        return ReadColor(color, alpha: 255);
    }

    static Color ReadColor(XElement element, byte alpha)
    {
        return Color.FromArgb(
            alpha,
            (byte)element.ReadDouble("r"),
            (byte)element.ReadDouble("g"),
            (byte)element.ReadDouble("b"));
    }

    #endregion
}
