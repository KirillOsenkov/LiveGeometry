#:property Nullable=disable
#:property PublishAot=false

// gallerize - turns old drawings into the drawings of the gallery
// (Main/Avalonia/LiveGeometry/Gallery/Drawings). Two sources: the Windows Phone app's .lgf
// (laid out for a portrait screen, with the text above and below the figure) and drawings of the
// DG 1.0 CD library, converted from .dgf by `LiveGeometry.Desktop.exe --check` and kept in
// Gallery/Sources. Here the window is landscape and of any size, so the text goes to the right
// of the figure, where it can never overlap it (a label is anchored at its top left corner and
// has a fixed size in pixels).
//
//   dotnet run tools/gallerize.cs -- <phone .lgf folder> <Gallery/Sources folder> <output folder>
//
// For each drawing listed below: drops the old heading and explanation labels, writes new ones
// (American school terminology, for a 13 year old), turns the grid off unless the drawing is
// about coordinates, unlocks the view and drops the phone's background. Everything else -
// figures, styles, coordinates - is kept as it was, apart from a `tweak` of the XML where a
// drawing needed one. The output is meant to be committed; run this again only to change all
// drawings at once.

using System.Globalization;
using System.Text;
using System.Xml.Linq;

if (args.Length < 3)
{
    Console.WriteLine("usage: gallerize <phone .lgf folder> <Gallery/Sources folder> <output folder>");
    return 1;
}

var drawings = new List<Sample>
{
    new("Bubbles", "Bubbles",
        "Grab any yellow point and the bubbles roll around, squeezing and growing - but they never overlap. How are they made?\n\nCheck \"Hint\" to see the trick.",
        remove: ["Label127", "Label128"]),
    new("InscribedCircle", "Inscribed Circle",
        "Every triangle has exactly one circle that touches all three sides: its incircle.\n\nThe green lines are the angle bisectors. They always meet in a single point, the incenter (gray) - the center of the incircle. A perpendicular from the incenter to a side gives the green point where the circle touches.\n\nDrag the yellow vertices: the circle always fits.",
        remove: ["Label28", "Label29"]),
    new("Morley", "Morley's Miracle",
        "Split each angle of a triangle into three equal parts. The trisectors next to each side meet in three points - and those points always form an equilateral triangle, no matter what triangle you start with!\n\nFrank Morley discovered this in 1899. It is so unexpected that it is known as Morley's miracle.\n\nDrag the yellow vertices and try to break it.",
        remove: ["Label26", "Label28"]),
    new("Pythagoras", "Pythagorean Theorem",
        "In a right triangle the square on the hypotenuse has the same area as the two squares on the legs together: a² + b² = c².\n\nRight now the legs are [AC] and [BC], and the hypotenuse is [ab].\nc² = [ab^2]\na² + b² = [AC^2+BC^2]\n\nDrag the yellow points - the two numbers always agree.",
        remove: ["Label39", "Label40", "Label41", "Label42", "Label43"],
        dependencies: ["A", "B", "C"]),
    new("SinScaleY", "Sine Wave: Amplitude",
        "y = [C.Y] · sin(x)\n\nDrag the yellow point up and down to change the number in front of the sine, the amplitude: how tall the wave is. What happens when it is negative? When it is zero?",
        remove: ["Label140"],
        dependencies: ["C"],
        grid: true),
    new("SinScaleX", "Sine Wave: Frequency",
        "y = sin([C.X] · x)\n\nDrag the yellow point left and right to change the number next to x, the frequency: how tightly the waves are packed. The height of the wave never changes.",
        remove: ["Label10"],
        dependencies: ["C"],
        grid: true),
    new("SquaresAroundRhombus", "Squares Around a Rhombus",
        "A rhombus is a quadrilateral with four equal sides. Build a square on each side, then connect the centers of the four squares.\n\nThe centers always form a square too!\n\nDrag the yellow points to change the rhombus.",
        remove: ["Label105", "Label106"]),
    new("CircumscribedCircle", "Circumscribed Circle",
        "Every triangle has exactly one circle through all three vertices: its circumcircle.\n\nThe gray points are the midpoints of the sides, and the gray lines are the perpendicular bisectors. They always meet in one point, the circumcenter - the center of the circle.\n\nDrag the yellow vertices. When does the center leave the triangle?",
        remove: ["Label1", "Label2", "Label46"]),
    new("CarpentersSquare", "Carpenter's Square",
        "A carpenter's square is a tool with a perfect right angle. Slide it so that its two arms always touch two fixed nails.\n\nWhat path does the corner follow? Drag it and guess, then check \"Hint\".",
        remove: ["Label309", "Label310"]),
    new("FibonacciSpiral", "Fibonacci Spiral",
        "1, 1, 2, 3, 5, 8, 13... Each Fibonacci number is the sum of the two before it.\n\nSquares with these side lengths fit together perfectly, and a quarter circle in each square makes a spiral. You can find it in seashells, sunflowers and pinecones.\n\nDrag the yellow points to turn and resize it.",
        remove: ["Label241"]),
    new("SquareBetweenSquares", "Square Between Squares",
        "Take any two squares (red and green). Connect their matching corners with four gray segments and mark the midpoints.\n\nThe midpoints form another square (yellow)! Its size is halfway between the two.\n\nDrag the yellow points to move and turn the squares.",
        remove: ["Label64", "Label65", "Label66"]),
    new("VanAubelsTheorem", "Van Aubel's Theorem",
        "Start with any quadrilateral and build a square on each side. Connect the centers of opposite squares.\n\nThe two segments are always perpendicular and have the same length - even when the quadrilateral is a complete mess.\n\nDrag the yellow vertices and see.",
        remove: ["Label116", "Label117"]),
    new("Desargues", "Desargues' Theorem",
        "Two triangles are \"in perspective from a point\" when the lines through matching vertices all meet in one point - like a shape and its shadow.\n\nThen the matching sides, extended, meet in three points that always lie on one line.\n\nDrag the vertices. Drawing by Chris Burrows.",
        remove: ["Label110", "Label111", "Label37"]),
    new("Bezier", "Bézier Curve",
        "A Bézier curve is controlled by four points: two endpoints and two handles. Each handle sets the direction in which the curve leaves its endpoint.\n\nEvery letter on this screen is drawn with Bézier curves. They are named after Pierre Bézier, who designed car bodies with them in the 1960s.\n\nDrag the points to bend the curve.",
        remove: ["Label85", "Label86", "Label87"]),
    new("TriangleOn3Lines", "Triangle on Three Lines",
        "Here are three parallel lines. Can you construct an equilateral triangle with one vertex on each line?\n\nDrag the yellow point: the triangle follows. Think about how it is done, then check \"Show solution\".",
        remove: ["Label77", "Label78"]),
    new("WireframeCube", "Wireframe Cube",
        "A cube you can turn - built from nothing but segments and midpoints.\n\nDrag the yellow points to rotate and stretch it.",
        remove: ["Label21", "Label22"]),
    new("SimsonLine", "Simson Line",
        "Triangle ABC is inscribed in a circle, and G is another point on the same circle. Drop perpendiculars from G to the three sides (extended if needed).\n\nThe three feet (green) always lie on one line: the Simson line, named after Robert Simson.\n\nDrag G around the circle and watch the line swing.",
        remove: ["Label242", "Label247"]),
    new("Ceva", "Ceva's Theorem",
        "Points D, E and F lie on the sides of triangle ABC. When do the segments AD, BE and CF pass through a single point?\n\nExactly when this product of ratios equals 1:\nAF/FB · BD/DC · CE/EA = [AF/FB*BD/DC*CE/EA]\n\nDrag D, E and F until the three segments meet, and watch the number.",
        remove: ["Label20", "Label21", "Label22"],
        dependencies: ["A", "B", "C", "D", "E", "F"]),
    new("AnglesInACircle", "Inscribed Angle",
        "A and C are fixed on the circle. Drag B along the circle: angle ABC doesn't change!\n\nAll inscribed angles that intercept the same arc are congruent, and each is half of the central angle AOC.\n\nNow move A and C so that AC is a diameter. The angle becomes exactly 90°.",
        remove: ["Label22", "Label23"]),
    new("SplittingTriangle", "Midsegments of a Triangle",
        "Connect the midpoints of the sides of a triangle. The three midsegments split it into four smaller triangles.\n\nAll four are congruent - the same shape and size - so each has exactly one quarter of the area, and sides half as long as the original.\n\nDrag the yellow vertices.",
        remove: ["Label103", "Label104"]),
    new("Parabola", "Parabola",
        "A parabola is the set of points that are equally far from a point (the focus F) and a line (the directrix AB).\n\nC runs along the line. E is where the perpendicular at C meets the perpendicular bisector of CF, so EC = EF. As C moves, E traces the parabola.\n\nDrag C, then move F closer to the line.",
        remove: ["Label348", "Label24"]),
    new("EllipseFromCircle", "Ellipse from a Circle",
        "Squash a circle and you get an ellipse.\n\nFrom a point on the circle drop a perpendicular to the axis and cut it in a fixed ratio. As the point runs around the circle, the cut point traces an ellipse.\n\nDrag the points to change how much the circle is squashed.",
        remove: ["Label14", "Label15", "Label16"]),
    new("Pappus", "Pappus's Theorem",
        "Pick three points on one line and three on another, and connect them crosswise.\n\nThe three places where the connecting segments cross always lie on one straight line. Pappus of Alexandria found this about 1,700 years ago.\n\nDrag any yellow point.",
        remove: ["Label54", "Label55"]),
    new("Trapezoid", "A Trapezoid Surprise",
        "ABDC is a trapezoid with bases AB and CD. E and H are the midpoints of the bases.\n\nDraw the diagonals and the lines from the midpoints, as shown. Lines through the crossing points cut the base AB into three equal parts: AL = LM = MB.\n\nDrag the vertices - it always works.",
        remove: ["Label73", "Label75", "Label92"]),
    new("QuadrilateralMidpoints", "Varignon's Theorem",
        "Connect the midpoints of the sides of any quadrilateral and you always get a parallelogram: opposite sides are equal and parallel, opposite angles are equal.\n\nIt even works when the quadrilateral crosses itself. Pierre Varignon proved it around 1700.\n\nDrag the yellow vertices.",
        remove: ["Label64", "Label65"]),
    new("SquareInSquare", "Square in a Square",
        "In square ABCD mark points E, F, G and H at the same distance from the corners: AE = BF = CG = DH.\n\nEFGH is a square too, and the four corner triangles are congruent right triangles. This picture is the heart of a famous proof of the Pythagorean theorem.\n\nDrag E along the side.",
        remove: ["Label73", "Label74"]),
    new("CompositionOfReflections", "Two Reflections",
        "Reflect a point across one line, then reflect the image across a second line.\n\nThe result is the same as one rotation around the point where the lines cross - by twice the angle between them.\n\nDrag the points and compare the two angles.",
        remove: ["Label80", "Label81"]),
    new("NapoleonsTheorem", "Napoleon's Theorem",
        "Build an equilateral triangle on each side of any triangle. Connect the centers of the three new triangles.\n\nYou always get another equilateral triangle! The theorem is named after Napoleon Bonaparte, who loved geometry - though nobody knows whether he really proved it.\n\nDrag the yellow vertices.",
        remove: ["Label198", "Label199", "Label200"]),
    new("ParabolaGraph", "Graph of a Parabola",
        "y = a·x² + b\na = [E.X]\nb = [F.X]\n\nThe two points under the x-axis are sliders. The blue one is a: it makes the parabola narrower, wider, or flips it upside down. The green one is b: it moves the parabola up and down.",
        remove: ["Label32", "Label33", "Label34"],
        dependencies: ["E", "F"],
        grid: true),
    new("ReuleauxTriangle", "Reuleaux Triangle",
        "A circle is not the only shape with the same width in every direction. This one is made of three 60° arcs, each centered at the opposite corner.\n\nRoll it between two parallel lines and they never move apart. That is why it can drill (almost) square holes, and why some coins have this shape.",
        remove: ["Label29", "Label30"]),
    new("CavalieriPrinciple", "Cavalieri's Principle",
        "The area of a triangle is ½ · base · height. Why doesn't it matter where the top is?\n\nThink of the triangle as a stack of thin slices. Slide the top sideways: every slice moves but keeps its length, so the area stays the same.\n\nDrag the top vertex.",
        remove: ["Label488", "Label489"]),
    new("CircleTangents", "Tangents to a Circle",
        "How do you draw the tangents from a point C to a circle with center A?\n\nFind the midpoint D of AC and draw the circle centered at D through C. It crosses the first circle at E and F - the points of tangency. It works because an angle inscribed in a semicircle is a right angle.\n\nDrag C.",
        remove: ["Label54", "Label55"]),
    new("Pentagon", "Regular Pentagon",
        "A regular pentagon built with compass and straightedge only:\n\n1. Draw radius AB and the radius AC perpendicular to it.\n2. D is the midpoint of AC.\n3. The circle centered at D through B meets line AC at E.\n4. BE is the side length: step it around the circle with the compass.\n\nDrag A and B.",
        remove: ["Label199", "Label200"]),

    // from the DG 1.0 CD library (Gallery/Sources); these were converted by today's reader,
    // so their intersections are right as they are - no IntersectionOrder="Legacy"
    new("Rose", "A Rose",
        "Every petal, leaf and the stem is a Bézier curve: a smooth curve steered by a few control points. Twenty-seven of them make this rose.\n\nThis is how fonts, logos and cartoon characters are drawn on computers.\n\nDrag any point and reshape the flower.",
        fromCD: true,
        remove: ["Label1", "Label2", "Label3"]),
    new("Castle", "Castle",
        "Mountains, a tree, a castle on a hill - all of it built from polygons and circles, just like a real drawing program.\n\nEvery corner is a point you can drag. Move the mountains, grow the tree, or give the castle another tower.",
        fromCD: true),
    new("Sierpinski", "Sierpinski Triangle",
        "Take a triangle, connect the midpoints of its sides, and you get four smaller triangles. Leave the middle one empty and repeat with the other three - again and again.\n\nThe result is a fractal: it looks the same however closely you zoom in. This one has five levels.\n\nDrag the three corners.",
        fromCD: true,
        remove: ["Label1"]),
    new("Spiral", "Spiral",
        "Point B is placed by a rule: its distance from the center is A's x-coordinate and its angle is A's y-coordinate. As A slides along the segment, B sweeps out a spiral.\n\nDrag the ends of the segment to wind the spiral tighter or looser.",
        fromCD: true,
        remove: ["Label1"]),
    new("SteinersProblem", "Steiner's Problem",
        "Four towns A, B, C and D want to build one shared airport E. Where should it go so that the total length of the four roads is as small as possible?\n\nRoads right now: [dist(E,A)+dist(E,B)+dist(E,C)+dist(E,D)]\n\nDrag E around and find the smallest number. Then look where E ended up - can you see the rule? Try a dented (non-convex) quadrilateral too.",
        fromCD: true,
        remove: ["Label1", "Label2", "Label3", "Label4", "Label5", "Label6", "Label7", "Label8"],
        dependencies: ["A", "B", "C", "D", "E"]),
    new("PickTheorem", "Pick's Theorem",
        "The corners of this polygon sit on grid points. Count the grid points inside it (I) and on its boundary (B). Then the area is\n\nArea = I + B/2 − 1\n\nRight now the area is [area(a,b,c,d,e,f,g,h)].\n\nTurn on Snap to grid and drag the corners to new grid points - the formula keeps working.",
        fromCD: true,
        remove: ["Label1", "Label2", "Label3", "Label4"],
        dependencies: ["A", "B", "C", "D", "E", "F", "G", "H"],
        grid: true,
        tweak: drawing => SetPlane(drawing, left: -6.5, bottom: -4.5, right: 5.5, top: 3.5)),
    new("EllipseEvolute", "Ellipse and Its Evolute",
        "Point L runs around an ellipse. At every position there is a circle that hugs the curve best, and M is its center.\n\nAs L travels, M traces the star-shaped curve inside: the evolute of the ellipse.\n\nDrag the yellow point to move L. Drag a and b to reshape the ellipse.",
        fromCD: true,
        remove: ["Label1", "Label2", "Label3", "Label4"],
        grid: true,
        tweak: drawing => SetPlane(drawing, left: -5.5, bottom: -4.6, right: 5.5, top: 3.6)),
    new("Icosahedron", "Icosahedron",
        "Twenty triangles, twelve corners, thirty edges: the icosahedron is the biggest of the five Platonic solids. A 20-sided die has this shape.\n\nThis is a flat drawing of it, so dragging a corner squashes the solid instead of turning it - try it.",
        fromCD: true,
        remove: ["Label1"]),
    new("Tetrahedron", "Tetrahedron",
        "The simplest solid: four corners, four triangular faces, six edges. It is a pyramid with a triangle for a base.\n\nDrag the corners to look at it from another side.",
        fromCD: true),
    new("BestFitCircle", "Best-Fit Circle",
        "Eight points that almost lie on a circle. Which circle fits them best?\n\nThis one is chosen so that the gaps between the points and the circle are as small as possible overall - the sum of their squares is the area of the square in the corner: [sqr(ai^2+bk^2+cm^2+od^2+qe^2+fs^2)].\n\nDrag the points. Move one far away and see how much the circle cares.",
        fromCD: true,
        remove: ["Label1", "Label2", "Label3"],
        dependencies: ["A", "I", "B", "K", "C", "M", "O", "D", "Q", "E", "F", "S"]),
    new("MeasuringDistance", "Measuring Across a Lake",
        "How far is it from A to B when a lake is in the way and you can't walk straight across?\n\nCheck \"Help 1\" and \"Help 2\" for the trick, and \"Show\" for the answer. It uses similar triangles - the same idea surveyors used before there were satellites.\n\nDrag A and B.",
        fromCD: true,
        remove: ["Label1", "Label2", "Label3", "Label4", "Label5", "Label6", "Label7"],
        tweak: drawing =>
        {
            // the check boxes were in a column of text far to the left; under the lake now
            var figures = drawing.Element("Figures");
            MoveFigure(figures, "ShowHideControl1", -5, -0.6);
            MoveFigure(figures, "ShowHideControl2", -3, -0.6);
            MoveFigure(figures, "ShowHideControl3", -1, -0.6);
        }),
    new("ComplexNumbers", "Complex Multiplication",
        "Think of the arrows as complex numbers: A = [A.X] + [A.Y]i and B = [B.X] + [B.Y]i.\n\nC is their product. Its angle is the sum of the two angles ([deg(ang(a,o,x))]° + [deg(ang(b,o,x))]° = [deg(ang(c,o,x))]°), and its length is the product of the two lengths ([AO] · [OB] = [OC]).\n\nMultiplying by a complex number rotates and stretches. Drag A and B.",
        fromCD: true,
        remove: ["Label1", "Label2", "Label3", "Label4", "Label5", "Label6", "Label7", "Label8"],
        dependencies: ["A", "B", "C", "O", "X"],
        grid: true,
        tweak: drawing => SetPlane(drawing, left: -2, bottom: -2.5, right: 6.5, top: 3.5)),
    new("Pascal", "Conic Through Five Points",
        "Any five points (no three on a line) determine exactly one conic section - an ellipse, a parabola or a hyperbola.\n\nThe red curve is that conic, drawn with a straightedge only: as M slides along the line, Pascal's theorem places the point A on the curve.\n\nDrag the five points and watch the ellipse become a hyperbola.",
        fromCD: true,
        remove: ["Label1", "Label2", "Label3"]),
    new("Ladder", "The Falling Ladder",
        "A ladder leans against a wall with a bucket of paint hanging on it. The ladder slips: its top slides down the wall, its foot slides out along the ground.\n\nWhat path does the bucket follow? Drag the top of the ladder down and see the curve.\n\nThen slide the bucket to the middle of the ladder. Now what shape is the path?",
        fromCD: true,
        remove: ["Label1", "Label2", "Label3", "Label4"],
        tweak: drawing =>
        {
            var figures = drawing.Element("Figures");

            // the ladder was a fifth of the house: five times longer, leaning at an angle, the
            // bucket near the top (the middle is what the text asks for next)
            SetParameter(figures, "Ladder length", -1.26);
            SetParameter(figures, "Drag me!", 0.3);
            SetParameter(figures, "The bucket", 0.2);

            // the slider for the length of the ladder ran across the house: above it
            MoveFigure(figures, "C", -2.9, 25.5);
            MoveFigure(figures, "D", 7.77, 25.5);

            // the curve the bucket draws, which the original asked the user to construct
            var bucket = figures.Elements("PointOnFigure").First(e => (string)e.Attribute("Name") == "The bucket");
            bucket.AddAfterSelf(new XElement("Locus",
                new XAttribute("Name", "BucketPath"),
                new XAttribute("Style", "GalleryLocus"),
                new XElement("Dependency", new XAttribute("Name", "The bucket")),
                new XElement("Dependency", new XAttribute("Name", "Drag me!"))));
        }),
};

var phoneSource = args[0];
var cdSource = args[1];
var output = args[2];
Directory.CreateDirectory(output);
foreach (var sample in drawings)
{
    var document = XDocument.Load(Path.Combine(sample.FromCD ? cdSource : phoneSource, sample.File + ".lgf"));
    Convert(document.Root, sample);
    var settings = new System.Xml.XmlWriterSettings() { Indent = true, Encoding = new UTF8Encoding(false), NewLineChars = "\r\n" };
    using (var writer = System.Xml.XmlWriter.Create(Path.Combine(output, sample.File + ".lgf"), settings))
    {
        document.Save(writer);
    }

    Console.WriteLine(sample.File);
}

return 0;

static void Convert(XElement drawing, Sample sample)
{
    var figures = drawing.Element("Figures");
    foreach (var name in sample.Remove)
    {
        var label = figures.Elements("Label").FirstOrDefault(e => (string)e.Attribute("Name") == name);
        if (label == null)
        {
            throw new Exception(sample.File + ": no label " + name);
        }

        label.Remove();
    }

    // labels that were left empty on the phone
    foreach (var empty in figures.Elements("Label").Where(e => string.IsNullOrWhiteSpace((string)e.Attribute("Text"))).ToArray())
    {
        empty.Remove();
    }

    // Since these drawings were made, Math.GetIntersectionOfCircleAndLine changed which of the
    // two intersections comes first when the line passes through the center of the circle
    // ("New code - preserves order"). That is how these drawings build their squares
    // (perpendicular at A, circle around A), and every square came out on the wrong side.
    // Whether a line passes through a center takes the geometry to tell (the center can be a
    // midpoint of two points of the line), so the reader does it: DrawingDeserializer, on
    // this attribute. Files don't record which version wrote them.
    if (!sample.FromCD)
    {
        drawing.SetAttributeValue("IntersectionOrder", "Legacy");
    }

    if (sample.Tweak != null)
    {
        sample.Tweak(drawing);
    }

    var viewport = drawing.Element("Viewport");
    viewport.SetAttributeValue("Grid", sample.Grid ? "true" : "false");
    viewport.SetAttributeValue("Locked", null);
    viewport.SetAttributeValue("Style", null);
    viewport.SetAttributeValue("Color", null);

    var styles = drawing.Element("Styles");
    styles.Elements("BackgroundStyle").Remove();
    styles.Add(TextStyle("GalleryTitle", fontSize: 30, color: "#FF1F4E8C", bold: true));
    styles.Add(TextStyle("GalleryText", fontSize: 16, color: "#FF2B3038", bold: false));
    styles.Add(new XElement("LineStyle",
        new XAttribute("Color", "#FFE0362B"),
        new XAttribute("StrokeWidth", "2.5"),
        new XAttribute("Name", "GalleryLocus")));

    // to the right of everything that has coordinates of its own
    var points = figures.Elements()
        .Where(e => Number(e, "X").HasValue && Number(e, "Y").HasValue)
        .Select(e => (X: Number(e, "X").Value, Y: Number(e, "Y").Value))
        .ToArray();
    double left = points.Min(p => p.X);
    double right = points.Max(p => p.X);
    double top = points.Max(p => p.Y);
    double bottom = points.Min(p => p.Y);
    double size = Math.Max(right - left, (top - bottom) * 1.6);
    double x = right + size * 0.12;
    double titleHeight = size * 0.075;

    figures.Add(Label("Title", "GalleryTitle", sample.Title, x, top, null));
    figures.Add(Label("Description", "GalleryText", Wrap(sample.Description, 46), x, top - titleHeight, sample.Dependencies));
}

/// <summary>Where a point on a figure sits along it</summary>
static void SetParameter(XElement figures, string pointName, double parameter)
{
    var point = figures.Elements("PointOnFigure").First(e => (string)e.Attribute("Name") == pointName);
    point.SetAttributeValue("Parameter", parameter.ToString(CultureInfo.InvariantCulture));
}

/// <summary>The part of the plane a drawing on the grid shows (see GalleryDrawing.GetPlane)</summary>
static void SetPlane(XElement drawing, double left, double bottom, double right, double top)
{
    var viewport = drawing.Element("Viewport");
    viewport.SetAttributeValue("Left", left.ToString(CultureInfo.InvariantCulture));
    viewport.SetAttributeValue("Bottom", bottom.ToString(CultureInfo.InvariantCulture));
    viewport.SetAttributeValue("Right", right.ToString(CultureInfo.InvariantCulture));
    viewport.SetAttributeValue("Top", top.ToString(CultureInfo.InvariantCulture));
}

static void MoveFigure(XElement figures, string name, double x, double y)
{
    var figure = figures.Elements().First(e => (string)e.Attribute("Name") == name);
    figure.SetAttributeValue("X", x.ToString(CultureInfo.InvariantCulture));
    figure.SetAttributeValue("Y", y.ToString(CultureInfo.InvariantCulture));
}

static double? Number(XElement element, string attribute)
{
    return double.TryParse((string)element.Attribute(attribute), NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ? value : null;
}

static XElement TextStyle(string name, double fontSize, string color, bool bold)
{
    return new XElement("TextStyle",
        new XAttribute("FontSize", fontSize.ToString(CultureInfo.InvariantCulture)),
        new XAttribute("Color", color),
        // where it isn't installed (the browser) TextStyle falls back to the app's own font;
        // naming that one ("Inter") doesn't work - it is embedded, not installed
        new XAttribute("FontFamily", "Segoe UI"),
        new XAttribute("Bold", bold ? "true" : "false"),
        new XAttribute("Italic", "false"),
        new XAttribute("Underline", "false"),
        new XAttribute("Name", name));
}

static XElement Label(string name, string style, string text, double x, double y, string[] dependencies)
{
    var label = new XElement("Label",
        new XAttribute("IsHitTestVisible", "false"),
        new XAttribute("Name", name),
        new XAttribute("Style", style),
        new XAttribute("Text", text.Replace("\n", @"\n")),
        new XAttribute("DecimalsToShow", "2"),
        new XAttribute("X", x.ToString("0.###", CultureInfo.InvariantCulture)),
        new XAttribute("Y", y.ToString("0.###", CultureInfo.InvariantCulture)));
    foreach (var dependency in dependencies ?? [])
    {
        label.Add(new XElement("Dependency", new XAttribute("Name", dependency)));
    }

    return label;
}

static string Wrap(string text, int width)
{
    var result = new StringBuilder();
    foreach (var paragraph in text.Split('\n'))
    {
        int column = 0;
        foreach (var word in paragraph.Split(' '))
        {
            if (column > 0 && column + 1 + word.Length > width)
            {
                result.Append('\n');
                column = 0;
            }
            else if (column > 0)
            {
                result.Append(' ');
                column++;
            }

            result.Append(word);
            column += word.Length;
        }

        result.Append('\n');
    }

    return result.ToString().TrimEnd('\n');
}

class Sample
{
    public Sample(
        string file,
        string title,
        string description,
        string[] remove = null,
        string[] dependencies = null,
        bool grid = false,
        bool fromCD = false,
        Action<XElement> tweak = null)
    {
        File = file;
        Title = title;
        Description = description;
        Remove = remove ?? [];
        Dependencies = dependencies;
        Grid = grid;
        FromCD = fromCD;
        Tweak = tweak;
    }

    public string File;
    public string Title;
    public string Description;
    public string[] Remove;
    public string[] Dependencies;
    public bool Grid;
    public bool FromCD;

    /// <summary>A change to the Drawing element before the text is added</summary>
    public Action<XElement> Tweak;
}
