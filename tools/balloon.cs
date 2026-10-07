#:property Nullable=disable
#:property PublishAot=false

// balloon - writes the "Pump Up the Balloon" gallery drawing: a floor pump on the grass with a
// hose to the knot of a balloon. The balloon is one closed Bezier path with automatic handles
// through 15 hidden points, two on each side of the neck and eleven around the body. The
// pump's handle slides on a hidden track, and how far it is pushed down is the air in the
// balloon: hidden points by coordinates work out from it how far the balloon lies over, how
// wide and tall it is and how deep its wrinkles are, and its points are expressions of those
// and of the knot, a free point. Two hidden paths are holes in it, the glints, and a hidden
// label is its tension: tight when it is nearly empty, so that the rubber creases, looser
// when it is full. Pushed all the way down, the handle pops it: points that have no value
// below that much air (the square root of a negative number) draw a burst, shreds and the
// torn neck, and the balloon's own points have none above it.
//
//   dotnet tools/balloon.cs -- <out.lgf>

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

if (args.Length < 1)
{
    Console.WriteLine("usage: balloon <out.lgf>");
    return 1;
}

// the air: 0 with the handle at the top of its track, 1 at the bottom
const double StartAir = 0.5;      // where the handle starts
const double FullFrom = 0.45;     // the air from which the balloon is no longer limp
const double OverfullFrom = 0.8;  // the air from which it bulges, the gauge's red zone
const double PopAt = 0.93;        // the air at which it pops

// where the knot starts, the other point to drag
const double KnotX = 0.25;
const double KnotY = -1.74;

// the size of the body (its half height when it stands): it grows with the square root of
// the air
const double EmptySize = 0.45;
const double SizeGrowth = 1.45;

// the empty balloon: lying over in a heap
const double Wrinkle = 0.6;       // the depth of the wrinkles when it is empty
const double WrinklePower = 1;    // how fast they fade as it fills
const double WrinkleAlong = 0.3;  // the part of a wrinkle along the balloon, the rest is across
const double Narrow = 0.58;       // how much narrower an empty balloon is
const double Stretch = 0;         // how much longer
const double Flop = 1.65;         // how far it lies over when empty, in radians
const double FlopAtLift = 1.42;   // how far it still lies when it starts to lift
const double LiftFrom = 0.2;      // the air at which it starts to stand up
const double LiftTo = 0.4;        // the air at which it stands
const double BendBase = 0.78;     // the part of the lean the neck takes; the top takes all of it
const double Squash = 0.55;       // how flat the side it lies on gets

// the neck: half its width at the knot and at its top, and its height, limp and full
const double NeckBottom = 0.04;
const double NeckHalf = 0.075;
const double NeckTopLimp = 0.16;
const double NeckTopFull = 0.11;

// the body: a teardrop, t from 0 at the neck round to 2 pi, half as wide as
// sin t |sin t/2|^power (see Tear): a pear when it stands up, pulled rounder as it fills
const int BodyPoints = 11;        // the points around the body, between the neck's
const double TearStart = 0.42;    // t of the first one, and 2 pi minus t of the last
const double PearPower = 0.45;
const double RoundPower = 0.3;
const double RoundGrowth = 1.5;   // how it pulls round on the way to the pop: faster and faster
const double RoundSquat = 0.05;   // how much less tall it is round
const double OverWide = 0.07;     // how much wider it is overfull

// how far each body point moves out (+) or in (-) when it is empty, times Wrinkle; the top
// one (Skin6) keeps its place
double[] wrinkles = { 0.35, -1, 0.5, 0.3, -0.25, 0, -0.3, 0.5, -0.75, 0.6, -0.5 };

// overfull it bulges off center, on the right: body point (from 0) -> how far across and
// along, in the egg's units
var bulges = new Dictionary<int, (double Across, double Along)>
{
    [1] = (0.02, 0),
    [2] = (0.07, 0.01),
    [3] = (0.07, 0.04),
    [4] = (0.02, 0.02),
};

// the tension of the automatic handles: tight when it is empty, looser when it is full
const double LimpTension = 1.5;
const double FullTension = 0.92;

// the glints: their middle above the egg's center, in its units, and the air from which
// they show
const double GlintCenter = 0.15;
const double ShinyFrom = 0.12;

// the gauge's needle, in degrees: where it points with no air, and how far it turns as the
// air goes to 1
const double NeedleEmpty = 210;
const double NeedleSweep = 240;

// where "POP!" is written, in pixels from the balloon's middle
const double PopOffsetX = -44;
const double PopOffsetY = -30;

// the view, and the scene that the gallery fits into the window
const double SceneLeft = -3.4;
const double SceneTop = 2.75;
const double SceneRight = 2.9;
const double SceneBottom = -2;

// the ends of the gradients, which all start at the top left of the box
const string ToBottom = "0,1";
const string ToRight = "1,0";
const string ToBottomRight = "1,1";

var figures = new StringBuilder();
var styles = new StringBuilder();

// the points written so far: what expressions name
var pointNames = new HashSet<string>();

// the sun: a soft glow, a halo and a disc, filled paths (a Circle would be drawn over the
// balloon's inside)
var sun = (X: 2.4, Y: 2.2);
Disc("SunGlow", sun, radius: 0.5);
Disc("SunHalo", sun, radius: 0.38);
Disc("Sun", sun, radius: 0.27);

Cloud("CloudA", (-1.45, 2.3), width: 1.25, height: 0.5);
Cloud("CloudB", (1.75, 1.15), width: 0.95, height: 0.38);

// the grass: a filled path far wider than any window, so that its ends never show, its top
// through five points on automatic handles, its sides and bottom straight
{
    const double End = 40;
    const double Bottom = -10;
    var top = new (double X, double Y)[] { (-End, -1.76), (-2.4, -1.72), (0.6, -1.78), (2.6, -1.74), (End, -1.76) };
    var anchors = new List<string>();
    for (int i = 0; i < top.Length; i++)
    {
        FixedPoint("Hill" + (i + 1), top[i].X, top[i].Y);
        anchors.Add("Hill" + (i + 1));
    }

    FixedPoint("HillRight", End, Bottom);
    FixedPoint("HillLeft", -End, Bottom);
    anchors.Add("HillRight");
    anchors.Add("HillLeft");
    BezierPath(
        "Grass",
        attributes: " Style=\"Grass\" Closed=\"true\" Filled=\"true\"",
        path: Repeat("C a a", top.Length - 1) + " L L L",
        sidesStyle: "NoEdge",
        dependencies: anchors);
}

// the floor pump: a base with two foot pegs, the barrel, the collar at its top and the
// nozzle the hose leaves from
Box("FootLeft", "PumpDark", (-3.2, -1.67, -2.88, -1.76));
Box("FootRight", "PumpDark", (-2.32, -1.67, -2.0, -1.76));
Box("PumpBase", "PumpDark", (-2.9, -1.54, -2.3, -1.76));
Box("Barrel", "PumpMetal", (-2.73, -0.62, -2.47, -1.58));
Box("Collar", "PumpDark", (-2.77, -0.56, -2.43, -0.67));
Box("Nozzle", "PumpDark", (-2.33, -1.58, -2.2, -1.68));

// the handle slides on a hidden track 1.3 long above the barrel, the rod from it into the
// barrel: its parameter, the part of the way from the top, is the air
FixedPoint("TrackTop", -2.6, 0.75);
FixedPoint("TrackBottom", -2.6, -0.55);
figures.AppendLine("    <Segment Name=\"Track\" Visible=\"false\">");
WriteDependencies(["TrackTop", "TrackBottom"]);
figures.AppendLine("    </Segment>");
figures.AppendLine($"    <PointOnFigure Name=\"Handle\" Style=\"Grip\" X=\"-2.6\" Y=\"{Format(0.75 - 1.3 * StartAir)}\" Parameter=\"{Format(StartAir)}\">");
WriteDependencies(["Track"]);
figures.AppendLine("    </PointOnFigure>");
pointNames.Add("Handle");
FixedPoint("RodFoot", -2.6, -0.56);
Point("GripLeft", "Handle.X - 0.17", "Handle.Y");
Point("GripRight", "Handle.X + 0.17", "Handle.Y");
Segment("Rod", "Handle", "RodFoot");
Segment("GripBar", "GripLeft", "GripRight");

// the knot comes after the handle: a tile's hover moves the first draggable point up and
// down evenly about its home (the second mostly down), and the handle's air has to go both
// ways, into the slump and toward full
figures.AppendLine($"    <FreePoint Name=\"Knot\" Style=\"KnotDot\" X=\"{Format(KnotX)}\" Y=\"{Format(KnotY)}\" />");
pointNames.Add("Knot");
FixedPoint("Outlet", -2.2, -1.63);

// the pressure gauge on the barrel: a dial with a red zone from OverfullFrom on
const double GaugeX = -2.6;
const double GaugeY = -1.12;
const double GaugeRadius = 0.2;
const double RedZoneRadius = 0.145;
const double NeedleLength = 0.16;
ShownPoint("GaugeCenter", Format(GaugeX), Format(GaugeY), style: "Hub");
FixedPoint("GaugeEdge", GaugeX + GaugeRadius, GaugeY);
figures.AppendLine("    <Circle Name=\"Gauge\" Style=\"GaugeFace\">");
WriteDependencies(["GaugeCenter", "GaugeEdge"]);
figures.AppendLine("    </Circle>");
FixedPoint("RedZoneStart", GaugeX + RedZoneRadius * Math.Cos(NeedleAngle(air: 1)), GaugeY + RedZoneRadius * Math.Sin(NeedleAngle(air: 1)));
FixedPoint("RedZoneEnd", GaugeX + RedZoneRadius * Math.Cos(NeedleAngle(air: OverfullFrom)), GaugeY + RedZoneRadius * Math.Sin(NeedleAngle(air: OverfullFrom)));
figures.AppendLine("    <CircleArc Name=\"RedZone\" Style=\"GaugeRed\">");
WriteDependencies(["GaugeCenter", "RedZoneStart", "RedZoneEnd"]);
figures.AppendLine("    </CircleArc>");

// the numbers the balloon is worked out from, two to a hidden point by coordinates; Air: X
// the air, Y the size of the body
var airExpression = "clamp((TrackTop.Y - Handle.Y) / 1.3, 0, 1)";
Point("Air", airExpression, $"{Format(EmptySize)} + {Format(SizeGrowth)} * sqrt({airExpression})");

// the needle, turned by the air
var needleAngle = $"{Format(NeedleAngle(air: 0))} - {Format(NeedleSweep * Math.PI / 180)} * Air.X";
Point("NeedleTip", $"{Format(GaugeX)} + {Format(NeedleLength)} * cos({needleAngle})", $"{Format(GaugeY)} + {Format(NeedleLength)} * sin({needleAngle})");
Segment("Needle", "GaugeCenter", "NeedleTip");

// Firm: X how round it pulls once it stands, from 0 at FullFrom to 1 at the pop, faster and
// faster, so that it grows taller by less and less; Y how overfull it is, from 0 at
// OverfullFrom to 1 at the pop
Point(
    "Firm",
    $"clamp((Air.X - {Format(FullFrom)}) / {Format(PopAt - FullFrom)}, 0, 1) ^ {Format(RoundGrowth)}",
    Smoothstep($"clamp((Air.X - {Format(OverfullFrom)}) / {Format(PopAt - OverfullFrom)}, 0, 1)"));

// Limp: X 1 when it is empty, 0 from FullFrom on; Y the depth of the wrinkles
Point(
    "Limp",
    $"max(0, 1 - Air.X / {Format(FullFrom)})",
    $"{Format(Wrinkle)} * max(0, 1 - Air.X / {Format(FullFrom)}) ^ {Format(WrinklePower)}");

// Body: the body's half width and half height
Point(
    "Body",
    $"Air.Y * (1 - {Format(Narrow)} * Limp.X) * (1 + {Format(OverWide)} * Firm.Y)",
    $"Air.Y * (1 + {Format(Stretch)} * Limp.X - {Format(RoundSquat)} * Firm.X)");

// Droop: X how far it lies over, in radians, until the air lifts it; Y the size of the
// glints, from a speck at ShinyFrom
var lift = $"clamp((Air.X - {Format(LiftFrom)}) / {Format(LiftTo - LiftFrom)}, 0, 1)";
var shine = $"clamp((Air.X - {Format(ShinyFrom)}) / 0.33, 0, 1)";
Point(
    "Droop",
    $"({Format(Flop)} - {Format(Flop - FlopAtLift)} * min(1, Air.X / {Format(LiftFrom)})) * (1 - {Smoothstep(lift)})",
    $"0.03 + 0.97 * {Smoothstep(shine)}");

// Sag: X how much of its width the side it lies on keeps; Y the height of the neck
Point(
    "Sag",
    $"1 - {Format(Squash)} * sin(Droop.X) ^ 6",
    $"{Format(NeckTopLimp)} - {Format(NeckTopLimp - NeckTopFull)} * Firm.X");

// gates, added to a point's X: 0 where the point is to show, no number where it is not (so
// that it doesn't exist): the balloon up to the pop, the burst after it (a hair later, so
// that the two never show together), the glints from ShinyFrom on
Point("PopBefore", $"0 * sqrt({Format(PopAt)} - Air.X)", "0");
Point("PopAfter", $"0 * sqrt(Air.X - {Format(PopAt + 1e-6)})", "0");
Point("Shiny", $"0 * sqrt(Air.X - {Format(ShinyFrom)})", "0");

// the middle of the body, where it bursts
Point("Ctr", "Knot.X", "Knot.Y + Sag.Y + Body.Y");

// the hose: a natural spline from the nozzle through a point that sags below the middle to
// the knot; it leaves the nozzle to the right and comes to the knot from below, on handles
// that are points, shorter when the ends are close
Point("HoseEnd", "Knot.X", "Knot.Y");
Point("HoseMid", "(Outlet.X + Knot.X) / 2", "(Outlet.Y + Knot.Y) / 2 - min(0.3, 0.08 + 0.08 * dist(Outlet, Knot))");
Point("HoseLeave", "Outlet.X + min(0.3, 0.15 * dist(Outlet, Knot))", "Outlet.Y");
Point("HoseArrive", "Knot.X", "Knot.Y - min(0.3, 0.15 * dist(Outlet, Knot))");
BezierPath(
    "Hose",
    attributes: " Smoothing=\"NaturalSpline\"",
    path: "C #3 a C a #4 C a a",
    sidesStyle: "HoseLine",
    dependencies: ["Outlet", "HoseMid", "HoseEnd", "HoseLeave", "HoseArrive"]);

// the balloon's points, counterclockwise from the knot: two on the right of the neck, the
// body's, two on the left of the neck; in the Dot style, which shows when the hint does
var skin = new List<string>();
skin.Add(NeckPoint("NeckBottomRight", across: NeckBottom, along: "0.02"));
skin.Add(NeckPoint("NeckTopRight", across: NeckHalf, along: "Sag.Y"));
for (int k = 0; k < BodyPoints; k++)
{
    double t = TearStart + (2 * Math.PI - 2 * TearStart) * k / (BodyPoints - 1);
    double pear = Tear(t, PearPower);
    double round = Tear(t, RoundPower);
    double up = -Math.Cos(t);
    var (bulgeAcross, bulgeAlong) = bulges.TryGetValue(k, out var bulge) ? bulge : (0, 0);
    double wrinkle = wrinkles[k];

    // across: the pear pulled round, bulging when overfull, wrinkled when limp, and on the
    // side it lies on (the right) flattened on the ground; along: the egg's height, wrinkled
    var width = $"({Format(pear)}{Term(round - pear, "Firm.X")}{Term(bulgeAcross, "Firm.Y")})";
    var flat = pear > 1e-6 ? " * Sag.X" : "";
    var across = $"(Body.X * {width} * (1{Term(wrinkle, "Limp.Y")}){flat})";
    var along = $"(Sag.Y + Body.Y * (1{Term(up, $"(1{Term(WrinkleAlong * wrinkle, "Limp.Y")})")}{Term(bulgeAlong, "Firm.Y")}))";
    double height = (1 + up) / 2;
    var name = "Skin" + (k + 1);
    Point(name, BalloonX(across, along, height, gate: "PopBefore"), BalloonY(across, along, height), style: "Dot");
    skin.Add(name);
}

skin.Add(NeckPoint("NeckTopLeft", across: -NeckHalf, along: "Sag.Y"));
skin.Add(NeckPoint("NeckBottomLeft", across: -NeckBottom, along: "0.02"));

// the glints, a crescent and a dot: places about GlintCenter as a radius and an angle in
// degrees, in the egg's units
(double Radius, double Degrees)[] crescentSpots = { (0.72, 100), (0.71, 124), (0.66, 148), (0.56, 133), (0.59, 111) };
var crescent = GlintPoints("Shine", crescentSpots.Select(spot => Polar(spot.Radius, spot.Degrees)).ToArray());
BezierPath(
    "Glint",
    attributes: " Visible=\"false\" Closed=\"true\"",
    path: Repeat("C a a", crescent.Count),
    sidesStyle: null,
    dependencies: crescent);
var (sparkleX, sparkleY) = Polar(radius: 0.56, degrees: 160);
const double SparkleRadius = 0.05;
var sparkle = GlintPoints(
    "Sparkle",
    Enumerable.Range(0, 3)
        .Select(corner => Math.PI / 2 + corner * 2 * Math.PI / 3)
        .Select(angle => (sparkleX + SparkleRadius * Math.Cos(angle), sparkleY + SparkleRadius * Math.Sin(angle)))
        .ToArray());
BezierPath(
    "Glint2",
    attributes: " Visible=\"false\" Closed=\"true\"",
    path: Repeat("C a a", sparkle.Count),
    sidesStyle: null,
    dependencies: sparkle);

// the tension of the balloon's automatic handles
figures.AppendLine($"    <Label Name=\"Tightness\" Visible=\"false\" Text=\"[1 + {Format(LimpTension - 1)} * Limp.X - {Format(1 - FullTension)} * Firm.X]\">");
WriteDependencies(["Limp", "Firm"]);
figures.AppendLine("    </Label>");

// the balloon: its points as anchors, the piece across the bottom of the neck straight; then
// the glints, which are holes, and the tension
var balloon = new List<string>(skin) { "Glint", "Glint2", "Tightness" };
BezierPath(
    "Balloon",
    attributes: $" Style=\"BalloonFill\" Closed=\"true\" Filled=\"true\" Tension=\"#{balloon.Count - 1}\"",
    path: Repeat("C a a", skin.Count - 1) + " L",
    sidesStyle: "BalloonRim",
    dependencies: balloon,
    handlesStyle: "BalloonHandle");

// the burst, after the pop: a star of seven spikes about the balloon's middle, sharp at the
// tips (handles on the anchor) and round in the valleys, growing with the air; and the word
// "POP!", the name of a point too small to see
{
    double[] tipRadii = { 1.25, 1.05, 1.3, 1.1, 1.22, 1.0, 1.28 };
    double[] valleyRadii = { 0.62, 0.55, 0.66, 0.58, 0.6, 0.64, 0.56 };
    double[] tipAngles = { 12, 62, 108, 160, 212, 262, 312 };
    var size = $"(0.85 + 3 * (Air.X - {Format(PopAt)}))";
    var burst = new List<string>();
    var pieces = new List<string>();
    for (int k = 0; k < tipAngles.Length; k++)
    {
        // the valley halfway between this tip and the one before
        double valleyAngle = (tipAngles[k] + (k == 0 ? tipAngles[^1] - 360 : tipAngles[k - 1])) / 2 * Math.PI / 180;
        double tipAngle = tipAngles[k] * Math.PI / 180;
        var valley = "BurstValley" + (k + 1);
        var tip = "BurstTip" + (k + 1);
        Point(valley, $"Ctr.X + {Format(valleyRadii[k] * Math.Cos(valleyAngle))} * {size} + PopAfter.X", $"Ctr.Y + {Format(valleyRadii[k] * Math.Sin(valleyAngle))} * {size}");
        Point(tip, $"Ctr.X + {Format(tipRadii[k] * Math.Cos(tipAngle))} * {size} + PopAfter.X", $"Ctr.Y + {Format(tipRadii[k] * Math.Sin(tipAngle))} * {size}");
        burst.Add(valley);
        burst.Add(tip);
        pieces.Add("C a 0,0");
        pieces.Add("C 0,0 a");
    }

    BezierPath(
        "Burst",
        attributes: " Style=\"BurstFill\" Closed=\"true\" Filled=\"true\"",
        path: string.Join(" ", pieces),
        sidesStyle: "BurstRim",
        dependencies: burst);
    ShownPoint("POP!", "Ctr.X + PopAfter.X", "Ctr.Y", style: "PopPoint");
    figures.AppendLine($"    <PointLabel Name=\"PopWord\" Style=\"PopText\" OffsetX=\"{Format(PopOffsetX)}\" OffsetY=\"{Format(PopOffsetY)}\" ShowName=\"true\" ShowCoordinates=\"false\">");
    WriteDependencies(["POP!"]);
    figures.AppendLine("    </PointLabel>");
}

// shreds of rubber flying out after the pop: crescents with sharp horns, on handles of their
// own; the middle of each flaps as it flies, and the curve follows it
{
    double[] directions = { 38, 136, 222, 322 };   // where each flies, in degrees
    double[] tilts = { 20, -35, 60, -15 };         // how it is turned, in degrees
    double[] flapRates = { 22, -26, 18, -24 };     // how fast its middle flaps
    double[] scales = { 1.25, 1.05, 1.15, 1.0 };

    // a shred about its own middle: a horn, the middle, the other horn; and the out and in
    // handle of each piece: horn to middle, middle to horn, the torn inner edge back
    (double X, double Y)[] shape = { (-0.12, -0.02), (0.01, 0.075), (0.12, -0.04) };
    (double X, double Y)[] handles = { (0.02, 0.06), (-0.06, 0), (0.06, 0), (-0.015, 0.06), (-0.045, 0.045), (0.045, 0.045) };
    var reach = $"(1.4 + 4 * (Air.X - {Format(PopAt)}))";
    for (int j = 0; j < directions.Length; j++)
    {
        double direction = directions[j] * Math.PI / 180;
        double tilt = tilts[j] * Math.PI / 180;
        double cosine = Math.Cos(tilt);
        double sine = Math.Sin(tilt);
        double scale = scales[j];
        var startX = $"Ctr.X + {Format(Math.Cos(direction))} * {reach}";
        var startY = $"Ctr.Y + {Format(Math.Sin(direction))} * {reach}";
        var anchors = new List<string>();
        for (int i = 0; i < shape.Length; i++)
        {
            double x = shape[i].X * scale;
            double y = shape[i].Y * scale;
            var name = $"Shred{j + 1}Corner{i + 1}";
            if (i == 1)
            {
                // the middle flaps: its height above the horns
                var flap = $"({Format(y)} + {Format(0.03 * scale)} * sin({Format(flapRates[j])} * (Air.X - {Format(PopAt)})))";
                Point(name, $"{startX}{PlusNumber(x * cosine)}{Term(-sine, flap)} + PopAfter.X", $"{startY}{PlusNumber(x * sine)}{Term(cosine, flap)}");
            }
            else
            {
                Point(name, $"{startX}{PlusNumber(x * cosine - y * sine)} + PopAfter.X", $"{startY}{PlusNumber(x * sine + y * cosine)}");
            }

            anchors.Add(name);
        }

        var turned = handles.Select(handle => HandleText(scale * (handle.X * cosine - handle.Y * sine), scale * (handle.X * sine + handle.Y * cosine))).ToArray();
        BezierPath(
            $"Shred{j + 1}",
            attributes: " Style=\"BalloonFill\" Closed=\"true\" Filled=\"true\"",
            path: $"C {turned[0]} {turned[1]} C {turned[2]} {turned[3]} C {turned[4]} {turned[5]}",
            sidesStyle: "ShredRim",
            dependencies: anchors);
    }
}

// what stays tied to the knot after the pop: the neck, torn into two curled flaps
{
    (double X, double Y)[] spots = { (0.04, 0.01), (0.16, 0.13), (-0.13, 0.17), (-0.04, 0.01) };
    var tatters = new List<string>();
    for (int i = 0; i < spots.Length; i++)
    {
        var name = "Tatter" + (i + 1);
        Point(name, $"Knot.X{PlusNumber(spots[i].X)} + PopAfter.X", $"Knot.Y{PlusNumber(spots[i].Y)}");
        tatters.Add(name);
    }

    BezierPath(
        "Tatters",
        attributes: " Style=\"BalloonFill\" Closed=\"true\" Filled=\"true\"",
        path: "C 0,0.08 -0.05,0.05 C -0.07,-0.09 0.05,-0.08 C 0.05,0.05 0,0.08 L",
        sidesStyle: "ShredRim",
        dependencies: tatters);
}

// the hint, under the explanation, and the box that shows it and the balloon's points
var hint =
    "The same 15 dots make the crumpled heap and the round balloon. Only where they sit changes, and how tightly the curve is pulled through them: tight when it is nearly empty, so the rubber creases, and looser when it is full, so it bulges round between them.\n\n" +
    "Once it stands up, each push adds the same amount of air, but the balloon grows taller by less and less: the new air has to spread around a bigger and bigger balloon.\n\n" +
    "Popped it? Pull the handle back up.";
figures.AppendLine($"    <Label Name=\"Hint\" Visible=\"false\" Style=\"GalleryText\" Text=\"{LabelText(hint)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"300\" WrapWidth=\"400\" Backdrop=\"true\" />");
figures.AppendLine("    <ShowHideControl Name=\"Hint box\" Style=\"GalleryText\" Show=\"false\" Text=\"Hint\" X=\"-3.15\" Y=\"2.4\">");
WriteDependencies([.. skin, "Hint"]);
figures.AppendLine("    </ShowHideControl>");

// the caption
var title = "Pump Up the Balloon";
var description =
    "Pull the red pump handle up to let the air out: the balloon flops over in a crumpled heap. Now push it down: the rubber lifts, stands up and stretches smooth, round and shiny. Keep an eye on the gauge! Drag the balloon anywhere and the hose follows.\n\n" +
    "The outline is one closed curve through 15 dots: check \"Hint\" to watch them spread apart. What happens if you push the handle all the way down?";
figures.AppendLine($"    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"{LabelText(title)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
figures.AppendLine($"    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"{LabelText(description)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");

// the styles: the drawing has a paper of its own, so every style looks the same under the
// dark theme (see Style)
Style("TextStyle", "GalleryTitle", "FontSize=\"30\" Color=\"#FF1A3D7C\" FontFamily=\"Segoe UI\" Bold=\"true\"");
Style("TextStyle", "GalleryText", "FontSize=\"16\" Color=\"#FF22324D\" FontFamily=\"Segoe UI\"");
Style("ShapeStyle", "SunGlow", "Fill=\"#33FFFFFF\" Color=\"#00000000\"");
Style("ShapeStyle", "SunHalo", "Fill=\"#59FFFFFF\" Color=\"#00000000\"");
Style("ShapeStyle", "Sun", "Color=\"#00000000\"", Gradient(ToBottomRight, ("#FFFFF59D", 0), ("#FFFFC107", 1)));
Style("LineStyle", "NoEdge", "Color=\"#00FFFFFF\"");
Style("ShapeStyle", "Cloud", "Color=\"#00000000\"", Gradient(ToBottom, ("#FFFFFFFF", 0), ("#FFEAF6FF", 1)));

// the grass is about 8.3 units deep, and its top 0.4 (to offset 0.05) is what the scene
// shows; below it, where a phone shows the caption over the grass, it fades to a pale green.
// Under the dark theme it stays a deeper green
var grassFill = Gradient(
    ToBottom,
    ("#FFA5DC78", 0),
    ("#FF7CC862", 0.006),
    ("#FF5BAF4C", 0.05),
    ("#FFD3ECBF", 0.16),
    ("#FFE8F5DC", 1));
var darkGrassFill = Gradient(
    ToBottom,
    ("#FFA5DC78", 0),
    ("#FF7CC862", 0.006),
    ("#FF5BAF4C", 0.05),
    ("#FF3F8E3C", 0.3),
    ("#FF2F7032", 1));
Style(
    "ShapeStyle",
    "Grass",
    "Color=\"#00000000\"",
    fill: grassFill,
    darkFill: darkGrassFill);

Style("ShapeStyle", "PumpMetal", "Color=\"#FF37474F\" StrokeWidth=\"1.5\"", Gradient(ToRight, ("#FF546E7A", 0), ("#FFB0BEC5", 0.35), ("#FF546E7A", 1)));
Style("ShapeStyle", "PumpDark", "Fill=\"#FF455A64\" Color=\"#FF37474F\"");
Style("LineStyle", "Rod", "Color=\"#FFB0BEC5\" StrokeWidth=\"5\"");
Style("LineStyle", "GripBar", "Color=\"#FFE53935\" StrokeWidth=\"9\"");
Style("PointStyle", "Grip", "Size=\"28\" Fill=\"#FFFF5252\" Color=\"#FFB71C1C\" StrokeWidth=\"2.5\"");
Style("PointStyle", "KnotDot", "Size=\"15\" Fill=\"#FF7F0000\" Color=\"#FF4A0000\" StrokeWidth=\"1.5\"");
Style("LineStyle", "HoseLine", "Color=\"#FF5C6BC0\" StrokeWidth=\"8\"");
Style("ShapeStyle", "GaugeFace", "Color=\"#FF263238\" StrokeWidth=\"2.5\"", Gradient(ToBottomRight, ("#FFFFFFFF", 0), ("#FFDDE6EA", 1)));
Style("LineStyle", "GaugeRed", "Color=\"#FFE53935\" StrokeWidth=\"4\"");
Style("LineStyle", "Needle", "Color=\"#FF263238\" StrokeWidth=\"2.5\"");
Style("PointStyle", "Hub", "Size=\"6\" Fill=\"#FF263238\" Color=\"#FF263238\"");
Style("ShapeStyle", "BalloonFill", "Color=\"#00000000\"", Gradient(ToBottomRight, ("#FFFF7C9C", 0), ("#FFF02848", 0.5), ("#FFC0001E", 1)));
Style("LineStyle", "BalloonRim", "Color=\"#FFB71C1C\" StrokeWidth=\"2\"");
Style("LineStyle", "ShredRim", "Color=\"#FFB71C1C\" StrokeWidth=\"1.5\"");
Style("PointStyle", "BalloonHandle", "Size=\"7\" Fill=\"#FFFFFFFF\" Color=\"#FFD50000\"");
Style("PointStyle", "Dot", "Size=\"8\" Fill=\"#FFFFFFFF\" Color=\"#FF8E0000\" StrokeWidth=\"1.5\"");
Style("ShapeStyle", "BurstFill", "Color=\"#00000000\"", Gradient(ToBottom, ("#FFFFF59D", 0), ("#FFFF9100", 1)));
Style("LineStyle", "BurstRim", "Color=\"#FFFF3D00\" StrokeWidth=\"3\"");
Style("PointStyle", "PopPoint", "Size=\"1\" Fill=\"#00FFFFFF\" Color=\"#00FFFFFF\" StrokeWidth=\"0.1\"");
Style("TextStyle", "PopText", "FontSize=\"40\" Color=\"#FFD50000\" FontFamily=\"Segoe UI\" Bold=\"true\"");

// the paper is the sky
var paper =Gradient(ToBottom, ("#FF9FDCFF", 0), ("#FFFFF6E0", 1));
var scene = $"Left=\"{Format(SceneLeft)}\" Top=\"{Format(SceneTop)}\" Right=\"{Format(SceneRight)}\" Bottom=\"{Format(SceneBottom)}\"";
var text = new StringBuilder();
text.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
text.AppendLine("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
text.AppendLine($"  <Viewport {scene}>");
text.AppendLine($"    <Background>{paper}</Background>");
text.AppendLine("  </Viewport>");
text.AppendLine($"  <Scene {scene} />");
text.AppendLine("  <Styles>");
text.Append(styles);
text.AppendLine("  </Styles>");
text.AppendLine("  <Figures>");
text.Append(figures);
text.AppendLine("  </Figures>");
text.AppendLine("</Drawing>");
File.WriteAllText(args[0], text.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
Console.WriteLine("wrote " + args[0]);
return 0;

// a hidden point by coordinates; x and y are expressions or numbers
void Point(string name, string x, string y, string style = null)
{
    WritePoint(name, " Visible=\"false\"" + StyleAttribute(style), x, y);
}

// one that shows
void ShownPoint(string name, string x, string y, string style)
{
    WritePoint(name, StyleAttribute(style), x, y);
}

// a hidden point that stays where it is
void FixedPoint(string name, double x, double y)
{
    Point(name, Format(x), Format(y));
}

// the points the expressions name are its dependencies
void WritePoint(string name, string attributes, string x, string y)
{
    var dependencies = DependenciesOf(x, y);
    var element = $"    <PointByCoordinates Name=\"{name}\"{attributes} X=\"{x}\" Y=\"{y}\"";
    if (dependencies.Count == 0)
    {
        figures.AppendLine(element + " />");
    }
    else
    {
        figures.AppendLine(element + ">");
        WriteDependencies(dependencies);
        figures.AppendLine("    </PointByCoordinates>");
    }

    pointNames.Add(name);
}

// the points the expressions name, in the order they name them
List<string> DependenciesOf(params string[] expressions)
{
    var result = new List<string>();
    foreach (var expression in expressions)
    {
        foreach (Match match in Regex.Matches(expression, "[A-Za-z_][A-Za-z0-9_]*"))
        {
            if (pointNames.Contains(match.Value) && !result.Contains(match.Value))
            {
                result.Add(match.Value);
            }
        }
    }

    return result;
}

void WriteDependencies(IEnumerable<string> dependencies)
{
    foreach (var dependency in dependencies)
    {
        figures.AppendLine($"      <Dependency Name=\"{dependency}\" />");
    }
}

// a segment in the style of its name
void Segment(string name, string start, string end)
{
    figures.AppendLine($"    <Segment Name=\"{name}\" Style=\"{name}\">");
    WriteDependencies([start, end]);
    figures.AppendLine("    </Segment>");
}

// a rectangle: a polygon on four fixed points named after its corners
void Box(string name, string style, (double Left, double Top, double Right, double Bottom) bounds)
{
    FixedPoint(name + "TopLeft", bounds.Left, bounds.Top);
    FixedPoint(name + "TopRight", bounds.Right, bounds.Top);
    FixedPoint(name + "BottomRight", bounds.Right, bounds.Bottom);
    FixedPoint(name + "BottomLeft", bounds.Left, bounds.Bottom);
    figures.AppendLine($"    <Polygon Name=\"{name}\" Style=\"{style}\">");
    WriteDependencies([name + "TopLeft", name + "TopRight", name + "BottomRight", name + "BottomLeft"]);
    figures.AppendLine("    </Polygon>");
}

// a Bezier path; path has a piece per anchor (L, or C and the two handles: an offset, a for
// automatic, #k for the dependency k), the dependencies are the anchors and then the points
// the path names, its holes and its tension
void BezierPath(
    string name,
    string attributes,
    string path,
    string sidesStyle,
    IEnumerable<string> dependencies,
    string handlesStyle = null)
{
    figures.AppendLine($"    <BezierPath Name=\"{name}\"{attributes} Path=\"{path}\">");
    if (sidesStyle != null)
    {
        figures.AppendLine($"      <Sides Style=\"{sidesStyle}\" />");
    }

    if (handlesStyle != null)
    {
        figures.AppendLine($"      <Handles Style=\"{handlesStyle}\" />");
    }

    WriteDependencies(dependencies);
    figures.AppendLine("    </BezierPath>");
}

// a filled path that looks like a circle, in the style of its name: three anchors a third of
// a turn apart, each handle along the tangent and 4/3 tan 30 degrees of the radius long (a
// cubic follows an arc of an angle a with handles 4/3 tan(a/4) long)
void Disc(string name, (double X, double Y) center, double radius)
{
    double handle = 4.0 / 3 * Math.Tan(Math.PI / 6) * radius;
    var anchors = new List<string>();
    for (int i = 0; i < 3; i++)
    {
        double angle = Math.PI / 2 + i * 2 * Math.PI / 3;
        FixedPoint(name + (i + 1), center.X + radius * Math.Cos(angle), center.Y + radius * Math.Sin(angle));
        anchors.Add(name + (i + 1));
    }

    var pieces = new List<string>();
    for (int i = 0; i < 3; i++)
    {
        double start = Math.PI / 2 + i * 2 * Math.PI / 3;
        double end = start + 2 * Math.PI / 3;

        // counterclockwise along the circle at the start, backwards at the end
        var outHandle = HandleText(-Math.Sin(start) * handle, Math.Cos(start) * handle);
        var inHandle = HandleText(Math.Sin(end) * handle, -Math.Cos(end) * handle);
        pieces.Add($"C {outHandle} {inHandle}");
    }

    BezierPath(
        name,
        attributes: $" Style=\"{name}\" Closed=\"true\" Filled=\"true\"",
        path: string.Join(" ", pieces),
        sidesStyle: "NoEdge",
        dependencies: anchors);
}

// a cloud: three puffs on a flat bottom, its anchors the valleys between the puffs and the
// corners of the bottom
void Cloud(string name, (double X, double Y) center, double width, double height)
{
    var corners = new (double X, double Y)[]
    {
        (center.X + 0.5 * width, center.Y - 0.3 * height),   // bottom right
        (center.X + 0.18 * width, center.Y + 0.32 * height), // valley right
        (center.X - 0.2 * width, center.Y + 0.22 * height),  // valley left
        (center.X - 0.5 * width, center.Y - 0.3 * height),   // bottom left
    };

    // how far each piece bulges out, by its length: the three puffs and the bottom
    double[] puffs = { 0.62, 0.66, 0.6, 0.08 };
    var anchors = new List<string>();
    for (int i = 0; i < corners.Length; i++)
    {
        FixedPoint(name + "Puff" + (i + 1), corners[i].X, corners[i].Y);
        anchors.Add(name + "Puff" + (i + 1));
    }

    var pieces = new List<string>();
    for (int i = 0; i < corners.Length; i++)
    {
        var start = corners[i];
        var end = corners[(i + 1) % corners.Length];
        var chord = (X: end.X - start.X, Y: end.Y - start.Y);
        double length = Math.Sqrt(chord.X * chord.X + chord.Y * chord.Y);

        // out of the cloud: to the right of the way round, which is counterclockwise
        double normalX = chord.Y / length;
        double normalY = -chord.X / length;
        double bulge = puffs[i] * length;
        var outHandle = HandleText(0.1 * chord.X + bulge * normalX, 0.1 * chord.Y + bulge * normalY);
        var inHandle = HandleText(-0.1 * chord.X + bulge * normalX, -0.1 * chord.Y + bulge * normalY);
        pieces.Add($"C {outHandle} {inHandle}");
    }

    BezierPath(
        name,
        attributes: " Style=\"Cloud\" Closed=\"true\" Filled=\"true\"",
        path: string.Join(" ", pieces),
        sidesStyle: "NoEdge",
        dependencies: anchors);
}

// a point of the neck, across from its axis and along it: it keeps its width and only leans
string NeckPoint(string name, double across, string along)
{
    Point(name, BalloonX(Format(across), along, height: 0, gate: "PopBefore"), BalloonY(Format(across), along, height: 0), style: "Dot");
    return name;
}

// the X of a point of the balloon, across from its axis and along it up from the knot,
// leaning over by the droop, more the further up (height from 0 at the neck to 1 at the
// top), so that an empty balloon lies on its side; and the gate added to it
string BalloonX(string across, string along, double height, string gate)
{
    return $"Knot.X + {across} * cos({Lean(height)}) + {along} * sin({Lean(height)}) + {gate}.X";
}

string BalloonY(string across, string along, double height)
{
    return $"Knot.Y - {across} * sin({Lean(height)}) + {along} * cos({Lean(height)})";
}

// the glints: holes in the balloon, in the egg's units, grown from a speck about their
// middle as the balloon fills (Droop.Y), leaning with the balloon as one piece, so that they
// keep their shape
List<string> GlintPoints(string prefix, (double X, double Y)[] spots)
{
    double middleX = spots.Average(spot => spot.X);
    double middleY = spots.Average(spot => spot.Y);
    var names = new List<string>();
    for (int i = 0; i < spots.Length; i++)
    {
        var unitX = $"({Format(middleX)}{Term(spots[i].X - middleX, "Droop.Y")})";
        var unitY = $"({Format(middleY)}{Term(spots[i].Y - middleY, "Droop.Y")})";
        var across = $"(Body.X * {unitX})";
        var along = $"(Sag.Y + Body.Y * (1 + {unitY}))";
        var name = prefix + (i + 1);
        Point(name, BalloonX(across, along, height: 0.8, gate: "Shiny") + " + PopBefore.X", BalloonY(across, along, height: 0.8));
        names.Add(name);
    }

    return names;
}

// a style with its colors repeated under the dark theme (darkFill, if it is given, instead
// of the fill)
void Style(
    string element,
    string name,
    string attributes,
    string fill = null,
    string darkFill = null)
{
    styles.AppendLine($"    <{element} Name=\"{name}\" {attributes}>");
    if (fill == null)
    {
        styles.AppendLine($"      <Dark {DarkAttributes(attributes)} />");
    }
    else
    {
        styles.AppendLine($"      <Fill>{fill}</Fill>");
        styles.AppendLine($"      <Dark {DarkAttributes(attributes)}>");
        styles.AppendLine($"        <Fill>{darkFill ?? fill}</Fill>");
        styles.AppendLine("      </Dark>");
    }

    styles.AppendLine($"    </{element}>");
}

// the color attributes of a style
static string DarkAttributes(string attributes)
{
    return string.Join(" ", Regex.Matches(attributes, "(Color|Fill)=\"[^\"]*\"").Select(match => match.Value));
}

static string Gradient(string endPoint, params (string Color, double Offset)[] stops)
{
    var sb = new StringBuilder();
    sb.Append($"<LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"{endPoint}\">");
    foreach (var stop in stops)
    {
        sb.Append($"<GradientStop Color=\"{stop.Color}\" Offset=\"{Format(stop.Offset)}\" />");
    }

    sb.Append("</LinearGradientBrush>");
    return sb.ToString();
}

static string StyleAttribute(string style) => style == null ? "" : $" Style=\"{style}\"";

// how far the balloon leans at a height: the neck BendBase of the droop, the top all of it
static string Lean(double height) => $"Droop.X * {Format(BendBase + (1 - BendBase) * height)}";

// the teardrop's half width at t: sin t, pinched toward the neck (t = 0) by |sin t/2|^power;
// the smaller the power, the rounder
static double Tear(double t, double power) => Math.Sin(t) * Math.Pow(Math.Abs(Math.Sin(t / 2)), power);

// a place on a glint: radius from its middle at the angle in degrees
static (double X, double Y) Polar(double radius, double degrees)
{
    double angle = degrees * Math.PI / 180;
    return (radius * Math.Cos(angle), GlintCenter + radius * Math.Sin(angle));
}

// where the needle points with that much air, in radians
static double NeedleAngle(double air) => (NeedleEmpty - NeedleSweep * air) * Math.PI / 180;

// a smooth step: 0 to 1 as value goes from 0 to 1, flat at both ends
static string Smoothstep(string value) => $"{value} ^ 2 * (3 - 2 * {value})";

// " + factor * expression" or " - ...", without the factor when it is 1, nothing when it is 0
static string Term(double factor, string expression)
{
    if (Math.Abs(factor) < 1e-9)
    {
        return "";
    }

    var sign = factor < 0 ? " - " : " + ";
    var size = Math.Abs(factor);
    return sign + (Math.Abs(size - 1) < 1e-9 ? expression : Format(size) + " * " + expression);
}

// " + value" or " - value", nothing when it is 0
static string PlusNumber(double value)
{
    if (Math.Abs(value) < 1e-9)
    {
        return "";
    }

    return (value < 0 ? " - " : " + ") + Format(Math.Abs(value));
}

static string HandleText(double x, double y) => Format(x) + "," + Format(y);

static string Repeat(string piece, int count) => string.Join(" ", Enumerable.Repeat(piece, count));

// a label's text as the file writes it: escaped for XML, a line break as \n
static string LabelText(string value)
{
    return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("\n", "\\n");
}

// a number with at most 6 decimals, and no negative zero
static string Format(double value)
{
    var rounded = Math.Round(value, digits: 6);
    if (rounded == 0)
    {
        rounded = 0;
    }

    return rounded.ToString("0.######", CultureInfo.InvariantCulture);
}
