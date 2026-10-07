#:property Nullable=disable
#:property PublishAot=false

// pokeblob - writes the "Poke the Blob" gallery drawing: a gummy blob with a face, sitting on a
// plate, and a gumball to push into it. The blob's outline is one closed BezierPath with
// automatic handles through 16 hidden points (the Dots box shows them). Gum, the gumball's
// shine, is the point to drag; the points are PointByCoordinates whose expressions follow it
// and Core, a hidden free point in the middle of the blob:
//
// - Poke is how far the ball's center is from the core and which way; Press, Mood, Squeeze and
//   Gulp work out of that how deep the ball presses, how surprised the face is, how the dots
//   crowd toward the ball, and how far the bubble has opened once the skin has closed over it.
// - Each point of the outline is the nearer of where the skin would be with no ball (a ripple,
//   the far side bulging) and where it runs around the ball, seen from the core, the corner
//   between the two rounded, and no lower than the flat bottom.
// - A swallowed ball sits in a bubble, a path that is a hole in the blob and in its gloss.
// - The eyes, the mouth and the cheeks are pushed away from the ball; the pupils watch it.
//
//   dotnet tools/pokeblob.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: pokeblob <out.lgf>");
    return 1;
}

// Every length below is in the blob's own units and multiplied by Scale on the way into the
// file: the gallery stops zooming in at 200 px a unit, so a bigger drawing is a bigger blob on
// a big screen
const double Scale = 1.25;

// the blob at rest: a circle around the core, sitting on a flat bottom
const double RestRadius = 1.4;
const double CoreY = 0.2;           // the core is at (0, CoreY), the plate and the face built on it
const double Floor = 1.25;          // the flat bottom, this far below the core
const double FloorSoftness = 0.08;  // how round the bottom's corners are

// the gumball; its center starts StartDistance from the core, StartAngle degrees
// counterclockwise from the right, already pressing a little into the skin
const double BallRadius = 0.22;
const double ShineOffset = 0.085;   // the shine, the point to drag, is this far left of the center and as far up
const double StartDistance = 1.36;
const double StartAngle = 48;

// the dent the ball presses into the skin
const double HugRadius = 0.212;     // where the skin runs around the ball, from its center: just under its rim (the ball is drawn over the skin)
const double WrapDegrees = 62;      // how far round the ball the skin hugs it
const double LipRoundness = 0.08;   // how round the lips of the dent are
const double Bulge = 0.1;           // the far side bulges by twice this times the press
const double Crowding = 0.7;        // how much the dots crowd toward the ball

// the gulp
const double DeepestPoke = 1.02;    // the ball's distance from the core at the deepest dent: nearer, the skin closes over the ball...
const double ClosingBand = 0.12;    // ...and has closed this much nearer
const double BubbleGrowth = 0.1;    // how much further in the bubble takes to open fully
const double BubbleMargin = 0.08;   // the open bubble's room around the ball

// the face
const double Clearance = 0.62;      // how near a swallowed ball comes to an eye or the mouth
const double CheekClearance = 0.47; // and to a cheek

var figures = new StringBuilder();

// the core: the middle of the blob at rest, a hidden free point that the blob, its face and
// the plate are built on
figures.AppendLine($"    <FreePoint Name=\"Core\" Visible=\"false\" X=\"0\" Y=\"{Length(CoreY)}\" />");

// the plate, its well and the blob's shadow, hugging the line the blob sits on
Oval("Plate", "Plate", (0, -Floor - 0.08), (2.15, 0.36));
Oval("Well", "Plate well", (0, -Floor - 0.06), (1.55, 0.22));
Oval("Shadow", "Shadow", (0.03, -Floor - 0.025), (1.18, 0.075));

// the gumball: Gum is its shine, the one point to drag, and Ball its center
double startAngle = StartAngle * Math.PI / 180;
double ballX = StartDistance * Math.Cos(startAngle) * Scale;
double ballY = (CoreY + StartDistance * Math.Sin(startAngle)) * Scale;
figures.AppendLine($"    <FreePoint Name=\"Gum\" Style=\"Shine\" X=\"{Format(ballX - ShineOffset * Scale)}\" Y=\"{Format(ballY + ShineOffset * Scale)}\" />");
Point("Ball", $"Gum.X + {Length(ShineOffset)}", $"Gum.Y - {Length(ShineOffset)}", ["Gum"]);
Point("GumRight", $"Ball.X + {Length(BallRadius)}", "Ball.Y", ["Ball"]);

// Poke.X: how far the ball is from the core; Poke.Y: which way
Point("Poke", "dist(Ball, Core)", "atan2(Ball.Y - Core.Y, Ball.X - Core.X)", ["Ball", "Core"]);

// Press.Y: 1 while the skin is open around the ball, 0 once it has closed over it; Press.X: how
// deep the ball presses, times that
var open = $"clamp((Poke.X - {Length(DeepestPoke - ClosingBand)}) / {Length(ClosingBand)}, 0, 1)";
Point("Press", $"clamp({Length(RestRadius + HugRadius)} - Poke.X, 0, {Length(1)}) * {open}", open, ["Poke"]);

// Mood.X: surprise, 0 to 1 (all surprise at a press of 0.45); Mood.Y: how high the ripple runs
var surprise = $"min(1, Press.X / {Length(0.45)})";
Point("Mood", surprise, $"{Length(0.025)} + {Length(0.05)} * {surprise}", ["Press"]);

// Squeeze.X: how much the dots crowd toward the ball (all of Crowding at a press of 0.12);
// Squeeze.Y: where the skin leaves the ball, as the distance from the core to the line it
// leaves along
double wrap = WrapDegrees * Math.PI / 180;
Point(
    "Squeeze",
    $"{Format(Crowding)} * min(1, Press.X / {Length(0.12)})",
    $"max({Length(0.01)}, Poke.X * {Format(Math.Cos(wrap))} - {Length(HugRadius)})",
    ["Press", "Poke"]);

// Gulp.X: how far the bubble has opened around the ball. Gulp exists only once the skin has
// closed over the ball (Gulp.Y has no value before that), and so do the bubble and what is
// built on it
var closed = Length(DeepestPoke - ClosingBand);
Point("Gulp", $"clamp(({closed} - Poke.X) / {Length(BubbleGrowth)}, 0, 1)", $"0 * sqrt({closed} - Poke.X)", ["Poke"]);

// The 16 dots of the outline, at rest every 22.5 degrees around the core. A dot's direction
// slides toward the ball's by Squeeze.X. SkinK.X is the dot's distance from the core with no
// ball (a ripple, the far side bulging); SkinK.Y where the skin runs around the ball, seen from
// the core: the near side of the circle of HugRadius around its center, then the line the skin
// leaves it along. RimK is at the nearer of the two while the skin is open around the ball (the
// corner between them rounded: the dent's lip), at SkinK.X once it has closed over the ball,
// and never below the flat bottom (that corner rounded too).
for (int k = 0; k < 16; k++)
{
    double restAngle = k * Math.PI / 8;
    var fromBall = k == 0 ? "-Poke.Y" : $"{Format(restAngle)} - Poke.Y";
    var slidFromBall = $"({fromBall} - Squeeze.X * sin({fromBall}))";
    var direction = k == 0 ? "(Squeeze.X * sin(Poke.Y))" : $"({Format(restAngle)} - Squeeze.X * sin({fromBall}))";
    var skin = $"{Length(RestRadius)} + Mood.Y * sin(3 * {direction} + Poke.Y) + {Format(Bulge)} * Press.X * (1 - cos({slidFromBall}))";
    var ballSide = $"Poke.X * cos({slidFromBall}) - sqrt(max(0, {Format(Math.Pow(HugRadius * Scale, 2))} - (Poke.X * sin({slidFromBall})) ^ 2))";
    var leaving = $"Squeeze.Y / max(0.05, cos({slidFromBall}) * {Format(Math.Cos(wrap))} - abs(sin({slidFromBall})) * {Format(Math.Sin(wrap))})";
    Point($"Skin{k}", skin, $"max({ballSide}, {leaving})", ["Poke", "Squeeze", "Mood", "Press"]);
    var gap = $"(Skin{k}.X - Skin{k}.Y)";
    var radius = $"(Skin{k}.X - Press.Y * ({gap} + sqrt({gap} ^ 2 + {Format(Math.Pow(LipRoundness * Scale, 2))})) / 2)";
    Point(
        $"Rim{k}",
        $"Core.X + {radius} * cos({direction})",
        $"Core.Y - {Length(Floor)} + {Length(FloorSoftness)} * log(1 + exp(({radius} * sin({direction}) + {Length(Floor)}) / {Length(FloorSoftness)}))",
        ["Core", $"Skin{k}", "Press", "Squeeze", "Poke"],
        style: "Dot");
}

// the face: where it rests, then pushed away from the ball, each part the way it goes when the
// ball is right on it
Point("EyeRestL", $"Core.X - {Length(0.45)}", $"Core.Y + {Length(0.38)}", ["Core"]);
Point("EyeRestR", $"Core.X + {Length(0.45)}", $"Core.Y + {Length(0.38)}", ["Core"]);
Point("MouthRest", "Core.X", $"Core.Y - {Length(0.3)}", ["Core"]);
(string Part, string AtRest, (double X, double Y) Away)[] features = [("EyeL", "EyeRestL", (-1, 0)), ("EyeR", "EyeRestR", (1, 0)), ("MouthC", "MouthRest", (0, -1))];
foreach (var (part, rest, away) in features)
{
    var (x, y) = Pushed(
        ($"{rest}.X", $"{rest}.Y"),
        $"dist(Ball, {rest})",
        Clearance,
        away,
        byPoke: true);
    Point(part, x, y, [rest, "Ball", "Press"]);
}

// the bubble the ball is swallowed in, with air in it: a hole in the blob and in its gloss. The
// first quarter of its wall, from BubbleE to BubbleN, shines brighter
var bubbleRadius = $"({Length(BallRadius)} + {Length(BubbleMargin)} * Gulp.X)";
Point("BubbleE", $"Ball.X + {bubbleRadius}", "Ball.Y", ["Ball", "Gulp"]);
Point("BubbleN", "Ball.X", $"Ball.Y + {bubbleRadius}", ["Ball", "Gulp"]);
Point("BubbleW", $"Ball.X - {bubbleRadius}", "Ball.Y", ["Ball", "Gulp"]);
Point("BubbleS", "Ball.X", $"Ball.Y - {bubbleRadius}", ["Ball", "Gulp"]);
BezierPath(
    "Bubble",
    "Air",
    "Bubble",
    ["BubbleE", "BubbleN", "BubbleW", "BubbleS"],
    parts: [("Piece1", "Bubble shine")]);

// the blob itself: the outline through the dots, its handles shown when a dot is clicked
var rims = Enumerable.Range(start: 0, count: 16).Select(k => $"Rim{k}").ToArray();
BezierPath(
    "Blob",
    "Gummy",
    "Gummy rim",
    rims,
    holes: ["Bubble"],
    handles: "Gummy handle");

// a gloss on the upper left, where the skin would be with no ball, under the face. When the
// ball dents the skin near it, it shrinks toward GlossEnd, its end away from the dent (to
// nothing when the dent is right on its middle)
double glossMiddle = 133 * Math.PI / 180;
double glossHalf = 25 * Math.PI / 180;
var glossAway = $"({Format(glossMiddle)} + {Format(glossHalf)} * clamp(atan2(sin({Format(glossMiddle)} - Poke.Y), cos({Format(glossMiddle)} - Poke.Y)) / {Format(15 * Math.PI / 180)}, -1, 1))";
Point(
    "GlossEnd",
    $"Core.X + 0.8 * {RestSkin(glossAway)} * cos({glossAway})",
    $"Core.Y + 0.8 * {RestSkin(glossAway)} * sin({glossAway})",
    ["Core", "Mood", "Poke", "Press"]);

// how much of the gloss is kept: less the harder the ball presses and the nearer the dent is to
// its middle
var glossKeep = $"(1 - min(1, Press.X / {Length(0.15)}) * exp(6 * (cos({Format(glossMiddle)} - Poke.Y) - 1))) ^ 1.5";

// the gloss's corners at rest: a direction in degrees, and how far out along it, as a part of
// the skin's distance
(string Name, double Degrees, double Fraction)[] gloss = [("GlossA", 108, 0.8), ("GlossB", 133, 0.885), ("GlossC", 158, 0.8), ("GlossD", 133, 0.785)];
foreach (var (name, degrees, fraction) in gloss)
{
    double angle = degrees * Math.PI / 180;
    Point(
        name,
        $"GlossEnd.X + {glossKeep} * (Core.X + {Format(fraction * Math.Cos(angle))} * {RestSkin(Format(angle))} - GlossEnd.X)",
        $"GlossEnd.Y + {glossKeep} * (Core.Y + {Format(fraction * Math.Sin(angle))} * {RestSkin(Format(angle))} - GlossEnd.Y)",
        ["GlossEnd", "Press", "Poke", "Core", "Mood"]);
}

BezierPath(
    "Gloss",
    "Gloss",
    "No line",
    gloss.Select(corner => corner.Name).ToArray(),
    holes: ["Bubble"]);

// blush under the eyes, kept off the ball too
foreach (var side in new[] { "L", "R" })
{
    var eye = "Eye" + side;
    var cheek = "Cheek" + side;
    int outward = side == "L" ? -1 : 1;
    var restX = $"{eye}.X{SignedLength(outward * 0.165)}";
    var restY = $"{eye}.Y - {Length(0.3)}";
    var (x, y) = Pushed(
        (restX, restY),
        $"sqrt(({restX} - Ball.X) ^ 2 + ({restY} - Ball.Y) ^ 2)",
        CheekClearance,
        (outward, 0),
        byPoke: false);
    Point(cheek, x, y, [eye, "Ball"]);
    Point(cheek + "In", $"{cheek}.X{SignedLength(-outward * 0.105)}", $"{cheek}.Y", [cheek]);
    Point(cheek + "Out", $"{cheek}.X{SignedLength(outward * 0.105)}", $"{cheek}.Y", [cheek]);
    string[] ends = side == "L" ? [cheek + "Out", cheek + "In"] : [cheek + "In", cheek + "Out"];
    BezierPath(
        cheek + "Blush",
        "Blush",
        "No line",
        ends,
        path: OvalPath(radiusY: 0.075));
}

// eyes that open wide as the blob is poked: an ellipse through the end of its across axis, so
// that it is drawn upright and its gradient runs top to bottom, then the top of the eye
foreach (var side in new[] { "L", "R" })
{
    var eye = "Eye" + side;
    Point(eye + "Side", $"{eye}.X + {Length(0.14)} + {Length(0.015)} * Mood.X", $"{eye}.Y", [eye, "Mood"]);
    Point(eye + "Top", $"{eye}.X", $"{eye}.Y + {Length(0.19)} + {Length(0.05)} * Mood.X", [eye, "Mood"]);
    Figure("Ellipse", eye + "White", "Eye white", [eye, eye + "Side", eye + "Top"]);
}

// the mouth: six points that go from a smile to an O as the face is surprised
(double X, double Y)[] smile = [(0.25, 0.03), (0.1, 0.01), (-0.1, 0.01), (-0.25, 0.03), (-0.12, -0.16), (0.12, -0.16)];
(double X, double Y)[] surprised = [(0.1, -0.07), (0.05, 0.045), (-0.05, 0.045), (-0.1, -0.07), (-0.05, -0.185), (0.05, -0.185)];
var lips = new List<string>();
for (int i = 0; i < smile.Length; i++)
{
    lips.Add($"Lip{i}");
    Point($"Lip{i}", $"MouthC.X{ByMood(smile[i].X, surprised[i].X)}", $"MouthC.Y{ByMood(smile[i].Y, surprised[i].Y)}", ["MouthC", "Mood"]);
}

BezierPath("Mouth", "Mouth", "Face line", lips.ToArray());

// pupils that watch the ball, and shrink as the face is surprised, with a sparkle each
foreach (var side in new[] { "L", "R" })
{
    var eye = "Eye" + side;
    var pupil = "Pupil" + side;
    Point(
        pupil,
        $"{eye}.X + {Length(0.06)} * (Ball.X - {eye}.X) / max({Length(0.05)}, dist(Ball, {eye}))",
        $"{eye}.Y + {Length(0.06)} * (Ball.Y - {eye}.Y) / max({Length(0.05)}, dist(Ball, {eye}))",
        [eye, "Ball"]);
    Point(pupil + "Edge", $"{pupil}.X + {Length(0.075)} - {Length(0.025)} * Mood.X", $"{pupil}.Y", [pupil, "Mood"]);
    Figure("Circle", pupil + "Disc", "Pupil", [pupil, pupil + "Edge"]);
    Point(
        "Spark" + side,
        $"{pupil}.X - 0.45 * ({pupil}Edge.X - {pupil}.X)",
        $"{pupil}.Y + 0.45 * ({pupil}Edge.X - {pupil}.X)",
        [pupil, pupil + "Edge"],
        style: "Sparkle",
        visible: true);
}

// once the skin has closed over the ball, the bubble's wall shines around it...
Figure("Circle", "BubbleFilm", "Bubble film", ["Ball", "BubbleE"]);

// the gumball, over everything but its glint and the points: the skin runs under its rim.
// (Nothing over it may be filled: a press there must take the ball, not what is built on the
// core too.)
Figure("Circle", "Gumball", "Gum", ["Ball", "GumRight"]);

// ...and in front of it, a glint that grows as the bubble opens
double glintMiddle = 135 * Math.PI / 180;
double glintHalf = 28 * Math.PI / 180;
Point(
    "GlintA",
    $"Ball.X + {Length(0.178)} * cos({Format(glintMiddle)} - {Format(glintHalf)} * Gulp.X)",
    $"Ball.Y + {Length(0.178)} * sin({Format(glintMiddle)} - {Format(glintHalf)} * Gulp.X)",
    ["Ball", "Gulp"]);
Point(
    "GlintB",
    $"Ball.X + {Length(0.178)} * cos({Format(glintMiddle)} + {Format(glintHalf)} * Gulp.X)",
    $"Ball.Y + {Length(0.178)} * sin({Format(glintMiddle)} + {Format(glintHalf)} * Gulp.X)",
    ["Ball", "Gulp"]);
Figure("CircleArc", "Glint", "Glint", ["Ball", "GlintA", "GlintB"]);

// the box that shows the dots, at the lower left
figures.AppendLine($"    <ShowHideControl Name=\"Dots\" Style=\"GalleryText\" Show=\"false\" Text=\"Dots\" X=\"{Length(-2.4)}\" Y=\"{Length(-1.42)}\">");
Dependencies(rims);
figures.AppendLine("    </ShowHideControl>");

// the caption, pinned to the screen (the gallery lays it out when the drawing opens)
const string Description =
    "This gummy blob's edge is one smooth curve through 16 hidden dots. Drag the pink gumball into it: nearby dots get pushed in, the rest bulge out, and the curve wraps around the ball with no corners. Check \"Dots\": they bunch up where the curve bends most. Click one to see the two handles steering the curve.\n\n" +
    "Push the gumball all the way in and... gulp! Now pull it back out (and watch that face).";
figures.AppendLine("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Poke the Blob\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
figures.AppendLine($"    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"{Escape(Description)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");

// The view and the scene take in the plate, the blob and the room above it for the gumball.
// The paper is the drawing's own under every theme, so the caption's styles keep their colors
// under Dark too.
var view = $"Left=\"{Length(-2.45)}\" Top=\"{Length(2.1)}\" Right=\"{Length(2.45)}\" Bottom=\"{Length(-1.65)}\"";
var text = $$"""
    <?xml version="1.0" encoding="utf-8"?>
    <Drawing Version="1" Creator="LiveGeometry.App">
      <Viewport {{view}}>
        <Background>
          <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
            <GradientStop Color="#FFFFF4E6" Offset="0" />
            <GradientStop Color="#FFE8E0FF" Offset="1" />
          </LinearGradientBrush>
        </Background>
      </Viewport>
      <Scene {{view}} />
      <Styles>
        <TextStyle Name="GalleryTitle" FontSize="30" Color="#FF4A2C6E" FontFamily="Segoe UI" Bold="true">
          <Dark Color="#FF4A2C6E" />
        </TextStyle>
        <TextStyle Name="GalleryText" FontSize="16" Color="#FF2E2A44" FontFamily="Segoe UI">
          <Dark Color="#FF2E2A44" />
        </TextStyle>
        <ShapeStyle Name="Plate" Color="#00000000">
          <Fill>
            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
              <GradientStop Color="#FFFFFFFF" Offset="0" />
              <GradientStop Color="#FFE9E5FA" Offset="0.8" />
              <GradientStop Color="#FFBDB3E8" Offset="1" />
            </LinearGradientBrush>
          </Fill>
        </ShapeStyle>
        <ShapeStyle Name="Plate well" Color="#00000000">
          <Fill>
            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
              <GradientStop Color="#FFE6E2F8" Offset="0" />
              <GradientStop Color="#FFFFFFFF" Offset="1" />
            </LinearGradientBrush>
          </Fill>
        </ShapeStyle>
        <LineStyle Name="No line" Color="#00000000" />
        <ShapeStyle Name="Shadow" Color="#00000000">
          <Fill>
            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
              <GradientStop Color="#304E3F8F" Offset="0" />
              <GradientStop Color="#664E3F8F" Offset="0.5" />
              <GradientStop Color="#304E3F8F" Offset="1" />
            </LinearGradientBrush>
          </Fill>
        </ShapeStyle>
        <ShapeStyle Name="Air" Fill="#FFE3FAF4" Color="#00000000" />
        <LineStyle Name="Bubble" Color="#CCFFFFFF" StrokeWidth="2" />
        <LineStyle Name="Bubble shine" Color="#FFFFFFFF" StrokeWidth="3.5" />
        <ShapeStyle Name="Gummy" Color="#00000000">
          <Fill>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
              <GradientStop Color="#C4C6FBDD" Offset="0" />
              <GradientStop Color="#C84DD0E1" Offset="0.5" />
              <GradientStop Color="#E000A68F" Offset="0.85" />
              <GradientStop Color="#EE00A68F" Offset="1" />
            </LinearGradientBrush>
          </Fill>
        </ShapeStyle>
        <LineStyle Name="Gummy rim" Color="#FF00897B" StrokeWidth="3" />
        <PointStyle Name="Gummy handle" Size="9" Fill="#FFFFFFFF" Color="#FF00897B" StrokeWidth="2" />
        <PointStyle Name="Dot" Size="11" Fill="#FFFFFFFF" Color="#FF00695C" StrokeWidth="2.5" />
        <ShapeStyle Name="Gloss" Fill="#A6FFFFFF" Color="#00000000" />
        <LineStyle Name="Face line" Color="#FF00574D" StrokeWidth="2.5" />
        <ShapeStyle Name="Eye white" Color="#FF00574D" StrokeWidth="2.5">
          <Fill>
            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
              <GradientStop Color="#FFFFFFFF" Offset="0.3" />
              <GradientStop Color="#FFE6E0F7" Offset="1" />
            </LinearGradientBrush>
          </Fill>
        </ShapeStyle>
        <ShapeStyle Name="Mouth" Fill="#FF0E3F3B" Color="#00000000" />
        <ShapeStyle Name="Pupil" Fill="#FF15292D" Color="#00000000" />
        <PointStyle Name="Sparkle" Size="6" Fill="#FFFFFFFF" Color="#00FFFFFF" StrokeWidth="0" />
        <ShapeStyle Name="Blush" Fill="#D9FF8AB5" Color="#00000000" />
        <ShapeStyle Name="Gum" Color="#FFB0144F" StrokeWidth="2">
          <Fill>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
              <GradientStop Color="#FFFFB3D9" Offset="0" />
              <GradientStop Color="#FFFF4F9A" Offset="0.55" />
              <GradientStop Color="#FFD81B60" Offset="1" />
            </LinearGradientBrush>
          </Fill>
        </ShapeStyle>
        <ShapeStyle Name="Bubble film" Color="#00FFFFFF" StrokeWidth="0">
          <Fill>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
              <GradientStop Color="#CCFFFFFF" Offset="0.15" />
              <GradientStop Color="#33FFFFFF" Offset="0.32" />
              <GradientStop Color="#00FFFFFF" Offset="0.45" />
              <GradientStop Color="#0000897B" Offset="0.6" />
              <GradientStop Color="#5900897B" Offset="0.85" />
            </LinearGradientBrush>
          </Fill>
        </ShapeStyle>
        <LineStyle Name="Glint" Color="#E6FFFFFF" StrokeWidth="3.5" />
        <PointStyle Name="Shine" Fill="#F2FFFFFF" Color="#00FFFFFF" StrokeWidth="0" />
      </Styles>
      <Figures>
    {{figures.ToString().TrimEnd()}}
      </Figures>
    </Drawing>
    """;
File.WriteAllText(args[0], text.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;

// a point given by expressions of its coordinates, built on the named figures; hidden unless
// it says otherwise
void Point(
    string name,
    string x,
    string y,
    string[] dependencies,
    string style = null,
    bool visible = false)
{
    var visibleText = visible ? "" : " Visible=\"false\"";
    var styleText = style != null ? $" Style=\"{style}\"" : "";
    figures.AppendLine($"    <PointByCoordinates Name=\"{name}\"{visibleText}{styleText} X=\"{x}\" Y=\"{y}\">");
    Dependencies(dependencies);
    figures.AppendLine("    </PointByCoordinates>");
}

// a closed, filled Bezier path through the anchors: automatic handles unless a path says
// otherwise, the holes other paths its inside leaves out, the parts pieces of the outline in a
// style of their own
void BezierPath(
    string name,
    string style,
    string sides,
    string[] anchors,
    string path = null,
    string[] holes = null,
    string handles = null,
    (string Name, string Style)[] parts = null)
{
    path ??= string.Join(" ", Enumerable.Repeat("C a a", anchors.Length));
    figures.AppendLine($"    <BezierPath Name=\"{name}\" Style=\"{style}\" Closed=\"true\" Filled=\"true\" Path=\"{path}\">");
    figures.AppendLine($"      <Sides Style=\"{sides}\" />");
    if (parts != null)
    {
        foreach (var part in parts)
        {
            figures.AppendLine($"      <Part Name=\"{part.Name}\" Style=\"{part.Style}\" />");
        }
    }

    if (handles != null)
    {
        figures.AppendLine($"      <Handles Style=\"{handles}\" />");
    }

    Dependencies(anchors);
    if (holes != null)
    {
        Dependencies(holes);
    }

    figures.AppendLine("    </BezierPath>");
}

// an oval under the blob, built on the core: two hidden points at the ends of its long axis
// and a path through them
void Oval(string name, string style, (double X, double Y) center, (double X, double Y) radius)
{
    Point(name + "Left", $"Core.X{SignedLength(center.X - radius.X)}", $"Core.Y{SignedLength(center.Y)}", ["Core"]);
    Point(name + "Right", $"Core.X{SignedLength(center.X + radius.X)}", $"Core.Y{SignedLength(center.Y)}", ["Core"]);
    BezierPath(
        name,
        style,
        "No line",
        [name + "Left", name + "Right"],
        path: OvalPath(radius.Y));
}

// a figure of the library on the named points, as it takes them: a circle (center, a point on
// it), an ellipse (center, the ends of its two axes), an arc (center, start, end)
void Figure(string type, string name, string style, string[] dependencies)
{
    figures.AppendLine($"    <{type} Name=\"{name}\" Style=\"{style}\">");
    Dependencies(dependencies);
    figures.AppendLine($"    </{type}>");
}

void Dependencies(IEnumerable<string> names)
{
    foreach (var name in names)
    {
        figures.AppendLine($"      <Dependency Name=\"{name}\" />");
    }
}

// the path of an oval through the ends of its long axis, the left one first: two cubics, over
// the top and under the bottom, whose handles go straight up or down by 4/3 of the half height
// (a cubic with both handles that high peaks at the half height)
static string OvalPath(double radiusY)
{
    var handle = Length(4.0 / 3 * radiusY);
    return $"C 0,{handle} 0,{handle} C 0,-{handle} 0,-{handle}";
}

// A part of the face pushed away from the ball: always kept at least `clearance` from its
// center (a swallowed ball, a deep poke), and with byPoke also pushed further the deeper the
// ball presses and the nearer it is. `away` is the way it goes when the ball is right on it.
static (string X, string Y) Pushed(
    (string X, string Y) rest,
    string distance,
    double clearance,
    (double X, double Y) away,
    bool byPoke)
{
    // the way from a place a little behind the ball's center, so that it is `away` when the
    // ball is right on the rest place
    var fromBallX = $"({rest.X} - Ball.X{SignedLength(0.03 * away.X)})";
    var fromBallY = $"({rest.Y} - Ball.Y{SignedLength(0.03 * away.Y)})";
    var length = $"sqrt({fromBallX} ^ 2 + {fromBallY} ^ 2)";

    // how much nearer than the clearance the ball is, or 0, the corner between the two rounded;
    // a poke pushes by 0.4 of the press, less the further the ball is
    var gap = $"{Length(clearance)} - {distance}";
    var keep = $"({gap} + sqrt(({gap}) ^ 2 + {Format(Math.Pow(0.06 * Scale, 2))})) / 2";
    var push = byPoke ? $"(0.4 * Press.X * exp(-({distance} / {Format(Scale)}) ^ 2) + {keep})" : $"({keep})";
    return ($"{rest.X} + {push} * {fromBallX} / {length}", $"{rest.Y} + {push} * {fromBallY} / {length}");
}

// the skin's distance from the core in a direction (an expression in radians), as the dots
// have it without the ball in the way and without their crowding
static string RestSkin(string angle)
{
    return $"({Length(RestRadius)} + Mood.Y * sin(3 * {angle} + Poke.Y) + {Format(Bulge)} * Press.X * (1 - cos({angle} - Poke.Y)))";
}

// a length added to an expression that goes from `calm` to `surprised` as Mood.X goes from 0 to 1
static string ByMood(double calm, double surprised)
{
    if (Math.Abs(surprised - calm) < 1e-9)
    {
        return SignedLength(calm);
    }

    return $"{SignedLength(calm)}{SignedLength(surprised - calm)} * Mood.X";
}

// a number as the file says it: at most five decimals, and never "-0"
static string Format(double value)
{
    var text = Math.Round(value, digits: 5).ToString("0.#####", CultureInfo.InvariantCulture);
    return text == "-0" ? "0" : text;
}

// a length in the blob's units, as the file says it
static string Length(double units)
{
    return Format(units * Scale);
}

// a length added to an expression: " + 0.3", " - 0.3", or nothing for 0
static string SignedLength(double units)
{
    double value = units * Scale;
    if (Math.Abs(value) < 1e-12)
    {
        return "";
    }

    return value > 0 ? $" + {Format(value)}" : $" - {Format(-value)}";
}

// text for an attribute: XML's escapes, and as a label's text keeps it, a backslash doubled
// and a line break as the two characters \n
static string Escape(string text)
{
    return text
        .Replace("&", "&amp;")
        .Replace("<", "&lt;")
        .Replace("\"", "&quot;")
        .Replace("\\", "\\\\")
        .Replace("\n", "\\n");
}
