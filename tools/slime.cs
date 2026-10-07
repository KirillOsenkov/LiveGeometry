#:property Nullable=disable
#:property PublishAot=false

// slime - writes the "Stretchy Slime" gallery drawing: galaxy slime held by two fists, to pull
// apart, push together and turn. The fists are free points drawn as emoji; the slime and all
// on it is worked out from them: points by coordinates whose X and Y are expressions (most of
// them hidden, only carrying two numbers for the expressions after them), and Bézier paths
// with automatic handles through such points:
//
// - Slime: one closed path through 24 anchors around two round blobs joined by a rope. The
//   further apart the fists, the longer the slime and the thinner its rope (its area stays
//   the same), and the more it sags. Past a length it grows a drip under its lowest point,
//   which lets go of a drop and grows again for the next one. Shade, a path on the same
//   anchors, darkens it below. The Dots box shows the anchors.
// - Bubble1..3: paths of four anchors that the slime leaves out as holes, stretched along it
//   as it gets longer, with a glint on each.
// - Gloss and Gleam: a crescent on each blob. Glitter1..4: sparkles in the slime.
// - Drop: a path of four anchors under the drip.
// - The stars on the paper: small polygons on hidden free points.
//
//   dotnet tools/slime.cs -- <out.lgf>

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

if (args.Length < 1)
{
    Console.WriteLine("usage: slime <out.lgf>");
    return 1;
}

// the fists start this far left and right of (0, 0)
const double FistStart = 1.25;

// the size of the fists' emoji, in pixels
const int FistSize = 64;

// the middles of the blobs sit this far inside the fists, so that the wrists stick out
const double Inset = 0.15;

// the slime's area, the same however it is pulled: pi e^2 + 2 k e for the blobs' radius e and
// k = 2 H (7/15 + 8 w / 15), with H half the distance between the blobs' middles and w the
// waist (the rope's half thickness in the middle over e)
const double Amount = 4.0;

// the waist w is WaistBulge / (1 + (H / (WaistLength / 2))^2), between WaistMin and 1: the
// longer the rope, the thinner it gets in the middle
const double WaistBulge = 1.05;
const double WaistLength = 2.8;
const double WaistMin = 0.22;

// how far the middle of the rope hangs down: SagRate times the square of the fists' distance,
// at most SagMax
const double SagMax = 0.9;
const double SagRate = 0.055;

// beyond this distance of the fists the slime stops stretching and stays between them
const double Cap = 7;

// the drip grows from a distance of DripStart, lets go of a drop at DropAt, springs back by
// SpringBack of itself and grows again for the next drop every Period further
const double DripStart = 2.55;
const double DropAt = 3.4;
const double Period = 0.74;
const double SpringBack = 0.55;

// how far the tip of a grown drip is below the surface
const double DripLength = 0.46;

// the neck's anchors are NeckWidth to the sides of the drip's middle when it has grown, and
// NeckFlare (1 - drip)^2 more while it grows: a short drip is a wide bell, a long one a
// teardrop on a narrow neck
const double NeckWidth = 0.065;
const double NeckFlare = 0.35;

// the bulb's anchors are BulbWidth to the sides when the drip has grown, BulbFlare less when
// it starts
const double BulbWidth = 0.13;
const double BulbFlare = 0.02;

// how far below the slime's underside the drop's middle starts (while the fists are level)
const double DropGap = 0.39;

// the slime's tension: 0.9 up to a distance of 2.2, then TensionSlope more for every unit
// further, at most TensionMax: the more it is pulled, the tauter its outline
const double TensionSlope = 0.06;
const double TensionMax = 1.15;

// the stars: the seed of their places, how many triangles, how far apart at least (|dx| + |dy|)
const int StarSeed = 11;
const int StarCount = 26;
const double StarSpacing = 0.6;

// the scene, and a narrower one for a phone in portrait. Dyadic numbers: the scene is kept as
// bottom and height, and Top must come back as written
const double SceneLeft = -2.75;
const double SceneTop = 1.625;
const double SceneBottom = -1.75;
const double PhoneLeft = -2.125;

// 1 while the fists are not crossed, -1 once they are (past 132 degrees), 0 at 120: what is
// across the rope (bubbles, glitter) goes over to the other side
const string Flip = "(1 - 2 * clamp(-3 * Axis.X - 1, 0, 1))";

var figures = new StringBuilder();

// the names of the figures written so far: a point by coordinates depends on those its
// expressions name
var known = new List<string>();

// ---- the stars, first: they are under everything else
WriteStars();

// ---- the fists, and the middle between them
figures.AppendLine($"    <FreePoint Name=\"FistL\" Style=\"LeftFist\" X=\"{Number(-FistStart)}\" Y=\"0\" />");
figures.AppendLine($"    <FreePoint Name=\"FistR\" Style=\"RightFist\" X=\"{Number(FistStart)}\" Y=\"0\" />");
known.Add("FistL");
known.Add("FistR");
WriteElement("MidPoint", "Middle", " Visible=\"false\"", new[] { "FistL", "FistR" });

// ---- the numbers the shape is made of, two on each hidden point

// Frame: X = the angle of the pair, Y = its length D, kept between 0.4 and Cap: beyond Cap the
// slime stays as it is, between the fists. Axis: the direction from the left fist to the right one
Hidden("Frame", "atan2(FistR.Y - FistL.Y, FistR.X - FistL.X)", $"clamp(dist(FistL, FistR), 0.4, {Number(Cap)})");
Hidden("Axis", "cos(Frame.X)", "sin(Frame.X)");

// Rope: X = H, half the distance between the middles of the two blobs; Y = w, the waist over
// the blobs' radius
string halfRope = $"Frame.Y / 2 - {Number(Inset)} * min(1, Frame.Y / 2.5)";
Hidden("Rope", halfRope, $"clamp({Number(WaistBulge)} / (1 + (({halfRope}) / {Number(WaistLength / 2)})^2), {Number(WaistMin)}, 1)");

// Body: X = e, the blobs' radius; Y = A, from the middle to a tip. The area stays the same:
// pi e^2 + 2 k e = Amount, with the half thickness e (1 - (1 - w)(1 - r^2)^2) along the rope
string ropeFactor = "(2 * Rope.X * (0.466667 + 0.533333 * Rope.Y))";
string blobRadius = $"(sqrt({ropeFactor}^2 + {Number(Math.PI * Amount)}) - {ropeFactor}) / pi";
Hidden("Body", blobRadius, $"Rope.X + {blobRadius}");

// Sag: X = how far the middle hangs down, less as the pair tilts; Y = how much of the drip
// there is for the tilt (all of it up to 10 degrees, none past 28, a smooth step between)
string level = "clamp((abs(Axis.X) - 0.88) / 0.105, 0, 1)";
Hidden("Sag", $"min({Number(SagMax)}, {Number(SagRate)} * Frame.Y^2) * abs(Axis.X)", $"(3 - 2 * {level}) * {level}^2");

// Drip: X = the drip, 0 to 1; it grows from DripStart to DropAt, lets a drop go, springs back
// and grows again for the next drop every Period. Y = how far into the period (the drop falls
// meanwhile)
string phase = $"((Frame.Y - {Number(DropAt)}) / {Number(Period)} - floor((Frame.Y - {Number(DropAt)}) / {Number(Period)}))";
Hidden("Drip", $"Sag.Y * clamp((Frame.Y - {Number(DripStart)}) / {Number(DropAt - DripStart)}, 0, 1) * (1 - {Number(SpringBack)} * clamp((Frame.Y - {Number(DropAt)}) * 200, 0, 1) * (1 - {phase}))", phase);

// Rib0..Rib5: X = the half thickness k sixths of the way from the middle to a tip, Y = how far
// the rope hangs down there. Along the rope (r = x / H up to 1) the half thickness is
// e (1 - (1 - w)(1 - r^2)^2) and the rope hangs (1 - r^2)^2 (1 + r^2) of the sag: both meet
// the blob with no corner. Past H, the blob's circle
Hidden("Rib0", "Body.X * Rope.Y", "Sag.X");
for (int i = 1; i <= 5; i++)
{
    var (half, hang) = Rib(Number(i / 6.0) + " * Body.Y");
    Hidden($"Rib{i}", half, hang);
}

// RibS: the rib under the shoulder of the drip side, which slides in from A/2 to A/3 as the
// drip grows, so that the long stretch between the shoulder and the neck follows the sagging rope
var (shoulderHalf, shoulderHang) = Rib(Shoulder("Drip.X"));
Hidden("RibS", shoulderHalf, shoulderHang);

// Low: X = where along the rope its lowest point is, when the pair is tilted (the drip hangs
// there); Y = how far the rope hangs there. Near the middle the rope hangs as a parabola:
// x = -H^2 sin / (2 sag), kept to A/6, halfway to the shoulders
string low = "clamp(-Rope.X^2 * Axis.Y / (2 * max(Sag.X, 0.05)), -0.166667 * Body.Y, 0.166667 * Body.Y)";
string lowShare = $"({low}) / Rope.X";
Hidden("Low", low, $"Sag.X * (1 - {lowShare}^2)^2 * (1 + {lowShare}^2)");

// ---- the outline: counterclockwise from the right tip, along the top from right to left,
// the left tip, along the bottom from left to right
var slimeAnchors = new List<string>();
slimeAnchors.Add(Tip("SlimeTipR", side: 1));
slimeAnchors.AddRange(Side(top: true));
slimeAnchors.Add(Tip("SlimeTipL", side: -1));
slimeAnchors.AddRange(Side(top: false));

// ---- the bubbles: an ellipse each, through four anchors at the ends of its axes
var bubbles = new[]
{
    new Bubble(Sixths: -2, Across: -0.2, Radius: 0.15),
    new Bubble(Sixths: 1, Across: 0.4, Radius: 0.065),
    new Bubble(Sixths: 2, Across: -0.45, Radius: 0.085),
};
var bubbleNames = new List<string>();
for (int index = 0; index < bubbles.Length; index++)
{
    var bubble = bubbles[index];
    string name = $"Bubble{index + 1}";
    string rib = $"Rib{Math.Abs(bubble.Sixths)}";
    string along = $"{Operand(bubble.Sixths / 6.0)} * Body.Y";

    // R: the semi-axes, stretched along and squeezed across by the same factor (the same
    // area), never longer than 2.5 times the width
    string stretch = "sqrt(clamp(Frame.Y / 2.2, 0.8, 2.5))";
    string semiAxisAcross = $"min({Number(bubble.Radius)} / {stretch}, 0.3 * {rib}.X)";
    Hidden($"{name}R", $"min({Number(bubble.Radius)} * {stretch}, 2.5 * {semiAxisAcross})", semiAxisAcross);

    // C: the middle, Across times the half thickness off the rope, further out when a fist is
    // near (0.45 from it at least), but inside 0.7 of the half thickness. With the fists
    // crossed every bubble goes to the other side, so the slime looks as it does uncrossed,
    // mirrored: the bubbles below keep to the dripping side's shoulders, the one above out of
    // the drip
    string fromFist = $"abs(Frame.Y / 2 - abs({along}))";
    string across = $"{Flip} * min(max({Number(Math.Abs(bubble.Across))} * {rib}.X, sqrt(max(0, 0.2025 - {fromFist}^2))), (0.62 + 0.08 * clamp((Frame.Y - 0.5) / 0.6, 0, 1)) * {rib}.X - {name}R.Y)";
    var (operatorX, operatorY) = AcrossOperators(bubble.Across);
    Hidden($"{name}C",
        $"Middle.X + {along} * Axis.X {operatorX} {across} * Axis.Y",
        $"Middle.Y + {along} * Axis.Y {operatorY} {across} * Axis.X - {rib}.Y");

    // T: the direction of the sagging rope there, which the long axis takes
    string share = $"clamp({along} / Rope.X, -1, 1)";
    string slope = $"(Sag.X * (2 * {share} + 4 * {share}^3 - 6 * {share}^5) / Rope.X)";
    string length = $"sqrt(1 + 2 * {slope} * Axis.Y + {slope}^2)";
    Hidden($"{name}T", $"Axis.X / {length}", $"(Axis.Y + {slope}) / {length}");

    var ring = new List<string>
    {
        Dot($"{name}E", $"{name}C.X + {name}R.X * {name}T.X", $"{name}C.Y + {name}R.X * {name}T.Y"),
        Dot($"{name}N", $"{name}C.X - {name}R.Y * {name}T.Y", $"{name}C.Y + {name}R.Y * {name}T.X"),
        Dot($"{name}W", $"{name}C.X - {name}R.X * {name}T.X", $"{name}C.Y - {name}R.X * {name}T.Y"),
        Dot($"{name}S", $"{name}C.X + {name}R.Y * {name}T.Y", $"{name}C.Y - {name}R.Y * {name}T.X"),
    };
    WritePath(
        name,
        style: "BubbleGlass",
        path: AutomaticHandles(anchors: 4),
        sides: "BubbleRim",
        dependencies: ring);
    bubbleNames.Add(name);
}

// ---- the slime, and its shade: the same outline, darker below. After the anchors come the
// bubbles, its holes, and the hidden label that says its tension
string tensionText = $"clamp(0.9 + {Number(TensionSlope)} * (Frame.Y - 2.2), 0.9, {Number(TensionMax)})";
WriteElement("Label", "Stretch", $" Visible=\"false\" Text=\"[{tensionText}]\" X=\"0\" Y=\"0\"", new[] { "Frame" });
var slimeDependencies = new List<string>(slimeAnchors);
slimeDependencies.AddRange(bubbleNames);
slimeDependencies.Add("Stretch");
string slimeTension = $"#{slimeDependencies.Count - 1}";
WritePath(
    "Slime",
    style: "SlimeFill",
    path: AutomaticHandles(slimeAnchors.Count),
    sides: "SlimeRim",
    dependencies: slimeDependencies,
    tension: slimeTension);
WritePath(
    "Shade",
    style: "SlimeShade",
    path: AutomaticHandles(slimeAnchors.Count),
    sides: "NoLine",
    dependencies: slimeDependencies,
    tension: slimeTension);

// ---- glints: on the upper left of each bubble, 0.7 of the way to its rim
foreach (var name in bubbleNames)
{
    string cos = $"(0.8 * {name}T.Y - 0.6 * {name}T.X)", sin = $"(0.8 * {name}T.X + 0.6 * {name}T.Y)";
    Shown($"{name}Glint",
        $"{name}C.X + 0.7 * ({name}R.X * {cos} * {name}T.X - {name}R.Y * {sin} * {name}T.Y)",
        $"{name}C.Y + 0.7 * ({name}R.X * {cos} * {name}T.Y + {name}R.Y * {sin} * {name}T.X)",
        style: "BubbleGlint");
}

// ---- gloss: a crescent on each blob, facing up (and out), never toward the rope. Shine:
// X = the angle for the left blob, of -Axis + (0.15, 0.85); Y = for the right one, of
// Axis + (-0.24, 0.72). Squeezed into a ball (stretched going to 0) there is no rope to keep
// away from: up and left, and up and right
string stretched = "clamp((Frame.Y - 0.5) / 1.7, 0, 1)";
Hidden("Shine",
    $"atan2({stretched} * (0.85 - Axis.Y) + 0.6 * (1 - {stretched}), {stretched} * (0.15 - Axis.X) - 0.6 * (1 - {stretched}))",
    $"atan2({stretched} * (Axis.Y + 0.72) + 0.6 * (1 - {stretched}), {stretched} * (Axis.X - 0.24) + 0.6 * (1 - {stretched}))");
WriteCrescent("Gloss", blob: -1, angle: "Shine.X", anchors: new (double Degrees, double Radius)[] { (42, 0.77), (0, 0.86), (-42, 0.77), (0, 0.64) });
WriteCrescent("Gleam", blob: 1, angle: "Shine.Y", anchors: new (double Degrees, double Radius)[] { (26, 0.78), (0, 0.86), (-26, 0.78), (0, 0.74) });

// ---- glitter: placed as the bubbles are, and going over with them when the fists cross
var glitter = new[]
{
    new Glitter(Sixths: -5, Across: -0.35, Style: "Glitter"),
    new Glitter(Sixths: -1, Across: 0.42, Style: "Glitter2"),
    new Glitter(Sixths: 3, Across: 0.5, Style: "Glitter2"),
    new Glitter(Sixths: 4, Across: -0.42, Style: "Glitter"),
};
for (int index = 0; index < glitter.Length; index++)
{
    var piece = glitter[index];
    string rib = $"Rib{Math.Abs(piece.Sixths)}";
    string along = Operand(piece.Sixths / 6.0);
    string across = $"{Number(Math.Abs(piece.Across))} * {Flip} * {rib}.X";
    var (operatorX, operatorY) = AcrossOperators(piece.Across);
    Shown($"Glitter{index + 1}",
        $"Middle.X + {along} * Body.Y * Axis.X {operatorX} {across} * Axis.Y",
        $"Middle.Y + {along} * Body.Y * Axis.Y {operatorY} {across} * Axis.X - {rib}.Y",
        style: piece.Style);
}

// ---- the drop: under the drip of whichever side hangs down, as big as the drip is for the
// tilt; it falls as the period goes on and the next one takes its place. 0 * sqrt(...) is 0
// where the root has a value and nothing elsewhere: there is no drop before DropAt, nor while
// the pair is tilted too far
string gate = $"0 * sqrt(Frame.Y - {Number(DropAt)}) + 0 * sqrt(Sag.Y - 0.15)";
Hidden("DropC",
    $"Middle.X + Low.X * Axis.X + sign(Axis.X) * Rib0.X * Axis.Y + {gate}",
    $"Middle.Y + Low.X * Axis.Y - Rib0.X * abs(Axis.X) - Low.Y - {Number(DropGap)} * Sag.Y - (0.15 * Drip.Y + 2.6 * Drip.Y^2)");
var drop = new List<string>
{
    Dot("DropTop", "DropC.X", "DropC.Y + 0.15 * Sag.Y"),
    Dot("DropLeft", "DropC.X - 0.1 * Sag.Y", "DropC.Y - 0.02 * Sag.Y"),
    Dot("DropBottom", "DropC.X", "DropC.Y - 0.12 * Sag.Y"),
    Dot("DropRight", "DropC.X + 0.1 * Sag.Y", "DropC.Y - 0.02 * Sag.Y"),
};

// pointed at the top: both handles of the top anchor are on it
WritePath(
    "Drop",
    style: "DropFill",
    path: "C 0,0 a C a a C a a C a 0,0",
    sides: "SlimeRim",
    dependencies: drop);

// ---- the Dots box, which shows the slime's anchors, and the caption
WriteElement("ShowHideControl", "DotsBox", " Style=\"GalleryText\" Show=\"false\" Text=\"Dots\" X=\"-2.05\" Y=\"-1.3\"", slimeAnchors);
string title = "Stretchy Slime";
string description = $"Grab the galaxy slime and pull! Drag the fists apart. There's only so much slime, so the longer it gets, the thinner it gets. Watch it sag, drip, and stretch its air bubbles into long ovals. Keep pulling: plop, plop! Push the fists together to squish it into a ball.\\n\\nIts edge is one smooth curve through {slimeAnchors.Count} dots, with no corners anywhere: a spline. Check Dots to watch the dots slide as you pull.";
figures.AppendLine($"    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"{title}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
figures.AppendLine($"    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"{description}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");

// ---- the file: a night sky for the paper, under both themes; so the caption's styles are
// light under both
var text = new StringBuilder();
text.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
text.AppendLine("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
text.AppendLine($"  <Viewport Left=\"{Number(SceneLeft)}\" Top=\"{Number(SceneTop)}\" Right=\"{Number(-SceneLeft)}\" Bottom=\"{Number(SceneBottom)}\">");
text.AppendLine("    <Background>");
text.AppendLine("      <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"0,1\">");
text.AppendLine("        <GradientStop Color=\"#FF140F3C\" Offset=\"0\" />");
text.AppendLine("        <GradientStop Color=\"#FF2A1A5E\" Offset=\"0.6\" />");
text.AppendLine("        <GradientStop Color=\"#FF3B1F6E\" Offset=\"1\" />");
text.AppendLine("      </LinearGradientBrush>");
text.AppendLine("    </Background>");
text.AppendLine("  </Viewport>");
text.AppendLine($"  <Scene Left=\"{Number(SceneLeft)}\" Top=\"{Number(SceneTop)}\" Right=\"{Number(-SceneLeft)}\" Bottom=\"{Number(SceneBottom)}\" />");
text.AppendLine($"  <Scene Left=\"{Number(PhoneLeft)}\" Top=\"{Number(SceneTop)}\" Right=\"{Number(-PhoneLeft)}\" Bottom=\"{Number(SceneBottom)}\" />");
text.AppendLine("  <Styles>");
text.AppendLine($"""
    <TextStyle Name="GalleryTitle" FontSize="30" Color="#FFFFD6F5" FontFamily="Segoe UI" Bold="true">
      <Dark Color="#FFFFD6F5" />
    </TextStyle>
    <TextStyle Name="GalleryText" FontSize="16" Color="#FFE8E0FF" FontFamily="Segoe UI">
      <Dark Color="#FFE8E0FF" />
    </TextStyle>
    <ShapeStyle Name="SlimeFill" Color="#00000000">
      <Fill>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
          <GradientStop Color="#FFFF6EC7" Offset="0" />
          <GradientStop Color="#FFB45CF5" Offset="0.3" />
          <GradientStop Color="#FF7C4DFF" Offset="0.55" />
          <GradientStop Color="#FF4D8DFF" Offset="0.78" />
          <GradientStop Color="#FF40E0FF" Offset="1" />
        </LinearGradientBrush>
      </Fill>
    </ShapeStyle>
    <ShapeStyle Name="SlimeShade" Color="#00000000">
      <Fill>
        <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
          <GradientStop Color="#00190A50" Offset="0" />
          <GradientStop Color="#00190A50" Offset="0.45" />
          <GradientStop Color="#40190A50" Offset="1" />
        </LinearGradientBrush>
      </Fill>
    </ShapeStyle>
    <LineStyle Name="SlimeRim" Color="#FFF3DDFF" StrokeWidth="2.5" />
    <LineStyle Name="NoLine" Color="#00FFFFFF" StrokeWidth="0.5" />
    <ShapeStyle Name="GlossFill" Color="#00000000">
      <Fill>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
          <GradientStop Color="#B3FFFFFF" Offset="0" />
          <GradientStop Color="#26FFFFFF" Offset="1" />
        </LinearGradientBrush>
      </Fill>
    </ShapeStyle>
    <ShapeStyle Name="BubbleGlass" Color="#00000000">
      <Fill>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
          <GradientStop Color="#B3FFFFFF" Offset="0" />
          <GradientStop Color="#40F0D8FF" Offset="0.5" />
          <GradientStop Color="#14FFFFFF" Offset="1" />
        </LinearGradientBrush>
      </Fill>
    </ShapeStyle>
    <PointStyle Name="BubbleGlint" Size="5" Fill="#F2FFFFFF" Color="#00FFFFFF" StrokeWidth="0.1" />
    <LineStyle Name="BubbleRim" Color="#CCFFFFFF" StrokeWidth="1.5" />
    <ShapeStyle Name="DropFill" Color="#00000000">
      <Fill>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
          <GradientStop Color="#FFE0A6FF" Offset="0" />
          <GradientStop Color="#FF7C4DFF" Offset="1" />
        </LinearGradientBrush>
      </Fill>
    </ShapeStyle>
    <PointStyle Name="Glitter" Shape="Diamond" Size="8" Fill="#F2FFFFFF" Color="#00FFFFFF" StrokeWidth="0.1" />
    <PointStyle Name="Glitter2" Shape="Diamond" Size="5" Fill="#E6FFC4EC" Color="#00FFFFFF" StrokeWidth="0.1" />
    <PointStyle Name="LeftFist" Character="🤜" Size="{FistSize}" Fill="#FFFFFFFF" />
    <PointStyle Name="RightFist" Character="🤛" Size="{FistSize}" Fill="#FFFFFFFF" />
    <PointStyle Name="Dot" Size="9" Fill="#FFFFE066" Color="#FF2A1460" StrokeWidth="1.5" />
    <PointStyle Name="SlimeHandle" Size="7" Fill="#FFFFFFFF" Color="#FF40E0FF" StrokeWidth="1.5" />
    <ShapeStyle Name="Star" Fill="#D9FFFFFF" Color="#D9FFFFFF" StrokeWidth="1.6" />
    <ShapeStyle Name="Star2" Fill="#8CFFFFFF" Color="#8CFFFFFF" StrokeWidth="1.2" />
    <ShapeStyle Name="Sparkle" Fill="#CCFFFFFF" Color="#00FFFFFF" StrokeWidth="0.5" />
""");
text.AppendLine("  </Styles>");
text.AppendLine("  <Figures>");
text.Append(figures);
text.AppendLine("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
Console.WriteLine("wrote " + args[0]);
return 0;

// the anchors of the top (from right to left) or the bottom (from left to right): three
// stations on the surface, the five anchors of the drip, three stations again
IEnumerable<string> Side(bool top)
{
    string side = top ? "Top" : "Bottom";
    int direction = top ? -1 : 1;
    string from = direction < 0 ? "R" : "L", to = direction < 0 ? "L" : "R";
    var result = new List<string>();
    for (int sixths = 5; sixths >= 3; sixths--)
    {
        result.Add(Station($"Slime{side}{from}{sixths}", -direction * sixths, top));
    }

    // the drip: neck, bulb, tip, bulb, neck; on the surface, evenly spaced, while there is no drip
    result.Add(DripAnchor($"Slime{side}Neck{from}", -direction * 2, top));
    result.Add(DripAnchor($"Slime{side}Bulb{from}", -direction, top));
    result.Add(DripAnchor($"Slime{side}Mid", sixths: 0, top));
    result.Add(DripAnchor($"Slime{side}Bulb{to}", direction, top));
    result.Add(DripAnchor($"Slime{side}Neck{to}", direction * 2, top));
    for (int sixths = 3; sixths <= 5; sixths++)
    {
        result.Add(Station($"Slime{side}{to}{sixths}", direction * sixths, top));
    }

    return result;
}

// a tip of the slime: 1 the right one, -1 the left one
string Tip(string name, int side)
{
    string sign = side > 0 ? "+" : "-";
    return Dot(name, $"Middle.X {sign} Body.Y * Axis.X", $"Middle.Y {sign} Body.Y * Axis.Y");
}

// an anchor on the surface, sixths sixths of the way from the middle to a tip
string Station(string name, int sixths, bool top)
{
    if (Math.Abs(sixths) == 3)
    {
        // on the side that hangs down (hanging = 1) the shoulder slides in with the drip and takes RibS
        string hanging = top ? "max(0, -sign(Axis.X))" : "max(0, sign(Axis.X))";
        string along = sixths < 0 ? $"(-{Shoulder($"Drip.X * {hanging}")})" : Shoulder($"Drip.X * {hanging}");
        var (shoulderX, shoulderY) = SurfaceAt(along, $"(Rib3.X + {hanging} * (RibS.X - Rib3.X))", $"(Rib3.Y + {hanging} * (RibS.Y - Rib3.Y + 0.05 * Drip.X))", top);
        return Dot(name, shoulderX, shoulderY);
    }

    var (x, y) = Surface(sixths, top);
    return Dot(name, x, y);
}

// The anchors that make the drip, on whichever side hangs down. Without a drip they are surface
// points; with one (drip up to 1) they go to a drip hanging straight down from the side's
// surface at the rope's lowest point: the necks beside it and a little down, the bulbs
// 0.68 DripLength down and to the sides, the tip DripLength down. Across the drip goes faster
// (blend) than down, so the neck is narrow early.
string DripAnchor(string name, int sixths, bool top)
{
    string drip = top ? "Drip.X * max(0, -sign(Axis.X))" : "Drip.X * max(0, sign(Axis.X))";
    string blend = $"min(1, 2.5 * {drip})";
    var (surfaceX, surfaceY) = Surface(sixths, top);
    var (middleX, middleY) = Surface(0, top);
    if (sixths == 0)
    {
        return Dot(name,
            $"{middleX} + {blend} * Low.X * Axis.X",
            $"{middleY} + {blend} * (Low.X * Axis.Y + Rib0.Y - Low.Y) - {Number(DripLength)} * {drip}");
    }

    // where the drip hangs from: the side's surface at the rope's lowest point
    string hangX = $"{middleX} + Low.X * Axis.X";
    string hangY = $"{middleY} + Low.X * Axis.Y + Rib0.Y - Low.Y";
    string sign = sixths < 0 ? "-" : "+";
    string targetX, targetY;

    // a short drip is a bell, its base (the necks) wide and hardly below the surface; a long one
    // a teardrop on a narrow neck: the neck pinches in as the drip grows, as a real drop's does
    if (Math.Abs(sixths) == 2)
    {
        string width = $"({Number(NeckWidth)} + {Number(NeckFlare)} * (1 - {drip})^2)";
        targetX = $"{hangX} {sign} {width} * Axis.X";
        targetY = $"{hangY} {sign} {width} * Axis.Y - {Number(DripLength)} * {drip} * (0.1 + 0.2 * {drip})";
    }
    else
    {
        string width = $"({Number(BulbWidth - BulbFlare)} + {Number(BulbFlare)} * {drip})";
        targetX = $"{hangX} {sign} {width} * sign(Axis.X)";
        targetY = $"{hangY} - {Number(0.68 * DripLength)} * {drip}";
    }

    return Dot(name,
        $"(1 - {blend}) * ({surfaceX}) + {blend} * ({targetX})",
        $"(1 - {blend}) * ({surfaceY}) + {blend} * ({targetY})");
}

// a point of the surface, sixths sixths of the way from the middle to a tip (negative: the left
// one), on the top or the bottom
static (string X, string Y) Surface(int sixths, bool top)
{
    int rib = Math.Abs(sixths);
    return SurfaceAt(sixths == 0 ? null : $"{Operand(sixths / 6.0)} * Body.Y", $"Rib{rib}.X", $"Rib{rib}.Y", top);
}

// a point of the surface, along the rope from the middle (null: the middle), with the half
// thickness and the hang there
static (string X, string Y) SurfaceAt(string along, string half, string hang, bool top)
{
    string x = along == null ? "" : $" + {along} * Axis.X";
    string y = along == null ? "" : $" + {along} * Axis.Y";
    return top
        ? ($"Middle.X{x} - {half} * Axis.Y", $"Middle.Y{y} + {half} * Axis.X - {hang}")
        : ($"Middle.X{x} + {half} * Axis.Y", $"Middle.Y{y} - {half} * Axis.X - {hang}");
}

// the half thickness and the hang at a distance along the rope (positive) from the middle
static (string X, string Y) Rib(string along)
{
    string share = $"min(1, {along} / Rope.X)";
    string thickness = $"(Body.X * (1 - (1 - Rope.Y) * (1 - {share}^2)^2))";
    string past = $"max(0, {along} - Rope.X)";
    return ($"sqrt(max(0, {thickness}^2 - {past}^2))", $"Sag.X * (1 - {share}^2)^2 * (1 + {share}^2)");
}

// where the shoulder station is along the rope for a drip: A/2 without a drip, near A/3 with one
static string Shoulder(string drip) => $"(0.5 * Body.Y - min(1, 1.6 * {drip}) * (0.166667 * Body.Y - {Number(NeckWidth / 2)}))";

// the operators that put a point across the rope, up for a positive across and down for a negative one
static (string X, string Y) AcrossOperators(double across) => across > 0 ? ("-", "+") : ("+", "-");

// A crescent around the middle of a blob (-1 the left one, 1 the right one), an anchor at the
// angle given plus each one's degrees, its radius times e out. The anchors at the two ends have
// their handles on them, so the crescent's ends are sharp.
void WriteCrescent(string name, int blob, string angle, (double Degrees, double Radius)[] anchors)
{
    string sign = blob < 0 ? "-" : "+";

    // in a ball (apart 0) around the middle and nearer the rim: the bubbles may be anywhere
    // around the fists, up to 0.62 e out
    string apart = "clamp((Frame.Y - 0.5) / 0.6, 0, 1)";
    var names = new List<string>();
    foreach (var (degrees, radius) in anchors)
    {
        string direction = degrees == 0 ? angle : $"{angle} {(degrees > 0 ? "+" : "-")} {Number(Math.Abs(degrees) * Math.PI / 180)}";
        names.Add(Hidden($"{name}{names.Count + 1}",
            $"Middle.X {sign} {apart} * Rope.X * Axis.X + ({Number(radius + 0.08)} - 0.08 * {apart}) * Body.X * cos({direction})",
            $"Middle.Y {sign} {apart} * Rope.X * Axis.Y + ({Number(radius + 0.08)} - 0.08 * {apart}) * Body.X * sin({direction})"));
    }

    WritePath(
        name,
        style: "GlossFill",
        path: "C 0,0 a C a 0,0 C 0,0 a C a 0,0",
        sides: "NoLine",
        dependencies: names);
}

// the stars: four-pointed sparkles where they look good, then small triangles at random places,
// apart from one another and from the sparkles, each turned at random
void WriteStars()
{
    var sparkles = new (double X, double Y, double Size)[] { (0.95, 1.4, 0.06), (-2.6, 0.2, 0.05), (2.5, -0.95, 0.055), (-1.4, -1.45, 0.045), (-3.2, -1.7, 0.04) };
    var random = new Random(StarSeed);
    var stars = new List<(double X, double Y, string Style)>();
    while (stars.Count < StarCount)
    {
        double x = -3.35 + 6.4 * random.NextDouble(), y = -2.1 + 3.95 * random.NextDouble();
        if (stars.Any(star => Math.Abs(star.X - x) + Math.Abs(star.Y - y) < StarSpacing)
            || sparkles.Any(sparkle => Math.Abs(sparkle.X - x) + Math.Abs(sparkle.Y - y) < StarSpacing))
        {
            continue;
        }

        stars.Add((Math.Round(x, digits: 2), Math.Round(y, digits: 2), random.NextDouble() < 0.4 ? "Star" : "Star2"));
    }

    int number = 0;
    foreach (var (x, y, size) in sparkles)
    {
        number++;
        var vertices = new List<string>();
        for (int i = 0; i < 8; i++)
        {
            double angle = Math.PI / 2 + i * Math.PI / 4;
            double radius = i % 2 == 0 ? size : size * 0.26;
            vertices.Add(WriteSkyPoint($"Sky{number}{(char)('a' + i)}", x + radius * Math.Cos(angle), y + radius * Math.Sin(angle)));
        }

        WritePolygon($"Sparkle{number}", "Sparkle", vertices);
    }

    foreach (var (x, y, style) in stars)
    {
        number++;
        double radius = style == "Star" ? 0.008 : 0.006;
        double turn = random.NextDouble() * Math.PI;
        var vertices = new List<string>();
        for (int i = 0; i < 3; i++)
        {
            double angle = turn + i * 2 * Math.PI / 3;
            vertices.Add(WriteSkyPoint($"Sky{number}{(char)('a' + i)}", x + radius * Math.Cos(angle), y + radius * Math.Sin(angle)));
        }

        WritePolygon($"Star{number}", style, vertices);
    }
}

string WriteSkyPoint(string name, double x, double y)
{
    figures.AppendLine($"    <FreePoint Name=\"{name}\" Visible=\"false\" X=\"{Number(x)}\" Y=\"{Number(y)}\" />");
    return name;
}

void WritePolygon(string name, string style, List<string> vertices)
{
    figures.AppendLine($"    <Polygon Name=\"{name}\" Style=\"{style}\">");
    foreach (var vertex in vertices)
    {
        figures.AppendLine($"      <Dependency Name=\"{vertex}\" />");
    }

    figures.AppendLine("    </Polygon>");
}

// a hidden point by coordinates
string Hidden(string name, string x, string y) => Point(name, x, y, " Visible=\"false\"");

// a hidden anchor of a path, a yellow dot when shown
string Dot(string name, string x, string y) => Point(name, x, y, " Visible=\"false\" Style=\"Dot\"");

// a point by coordinates that shows
string Shown(string name, string x, string y, string style) => Point(name, x, y, $" Style=\"{style}\"");

// a point by coordinates: it depends on the figures its expressions name
string Point(string name, string x, string y, string attributes)
{
    WriteElement("PointByCoordinates", name, $"{attributes} X=\"{x}\" Y=\"{y}\"", NamesIn(x + " " + y));
    return name;
}

// a figure with its dependencies; the expressions after it can name it
void WriteElement(string element, string name, string attributes, IEnumerable<string> dependencies)
{
    figures.AppendLine($"    <{element} Name=\"{name}\"{attributes}>");
    foreach (var dependency in dependencies)
    {
        figures.AppendLine($"      <Dependency Name=\"{dependency}\" />");
    }

    figures.AppendLine($"    </{element}>");
    known.Add(name);
}

// a closed and filled Bézier path: the anchors are its first dependencies, the paths after them
// its holes, and the figure that says its tension, if any, the last one
void WritePath(
    string name,
    string style,
    string path,
    string sides,
    IEnumerable<string> dependencies,
    string tension = null)
{
    string tensionAttribute = tension == null ? "" : $" Tension=\"{tension}\"";
    figures.AppendLine($"    <BezierPath Name=\"{name}\" Style=\"{style}\" Closed=\"true\" Filled=\"true\"{tensionAttribute} Path=\"{path}\">");
    figures.AppendLine($"      <Sides Style=\"{sides}\" />");
    figures.AppendLine("      <Handles Style=\"SlimeHandle\" />");
    foreach (var dependency in dependencies)
    {
        figures.AppendLine($"      <Dependency Name=\"{dependency}\" />");
    }

    figures.AppendLine("    </BezierPath>");
    known.Add(name);
}

// the figures an expression names, in the order it first names them
List<string> NamesIn(string expression)
{
    var result = new List<string>();
    foreach (Match match in Regex.Matches(expression, "[A-Za-z_][A-Za-z0-9_]*"))
    {
        if (known.Contains(match.Value) && !result.Contains(match.Value))
        {
            result.Add(match.Value);
        }
    }

    return result;
}

// a path whose handles are all automatic
static string AutomaticHandles(int anchors) => string.Join(" ", Enumerable.Repeat("C a a", anchors));

static string Number(double value) => Math.Round(value, digits: 6).ToString("0.######", CultureInfo.InvariantCulture);

// a number to follow an operator: a negative one in parentheses
static string Operand(double value) => value < 0 ? $"({Number(value)})" : Number(value);

// a bubble: Sixths sixths of the way from the middle to a tip (negative: the left one), Across
// times the half thickness there above the rope (negative: below), and the radius of a circle
// of its area
record Bubble(int Sixths, double Across, double Radius);

// a piece of glitter, placed as a bubble is
record Glitter(int Sixths, double Across, string Style);
