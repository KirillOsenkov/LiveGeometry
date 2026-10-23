using System.Linq;
using Avalonia.Controls.Shapes;

namespace LiveGeometry;

public class ReflectedPoint : PointBase, IPoint
{
    [PropertyGridVisible]
    [PropertyGridName("Reflection Of ")]
    public IFigure Source => Dependencies.ElementAt(0);

    [PropertyGridVisible]
    [PropertyGridName("Reflected About ")]
    public IFigure Mirror => Dependencies.ElementAt(1);

    /// <summary>A reflection in a circle is an inversion</summary>
    protected override string Kind => Dependencies.Count > 1 && Mirror is ICircle ? "Inverted point" : "Reflected point";

    /// <summary>"of A across line g", "of A through O", "of A in circle k"</summary>
    public override string Construction => Dependencies.Count < 2 ? null : "of " + ConstructionText.Of(Source) + MirrorText;

    /// <summary>The mirror as the construction says it, with its preposition: " across line g", " through O", " in circle k"</summary>
    public string MirrorText
    {
        get
        {
            var mirror = Mirror;
            var preposition = mirror is IPoint ? " through " : mirror is ICircle ? " in " : " across ";
            return preposition + ConstructionText.Of(mirror);
        }
    }

    protected override Shape CreateShape()
    {
        return Factory.CreateDependentPointShape();
    }

    protected override void OnDependenciesChanged()
    {
        mirrorPoint = Mirror as IPoint;
        mirrorLine = Mirror as ILine;
        mirrorCircle = Mirror as ICircle;
    }

    private IPoint mirrorPoint;
    private ILine mirrorLine;
    private ICircle mirrorCircle;

    public override void Recalculate()
    {
        var source = Point(0);
        if (mirrorPoint != null)
        {
            Coordinates = Math.GetSymmetricPoint(source, mirrorPoint.Coordinates);
        }
        else if (mirrorLine != null)
        {
            Coordinates = Math.GetSymmetricPoint(source, mirrorLine.Coordinates);
        }
        else if (mirrorCircle != null)
        {
            Coordinates = Math.GetSymmetricPoint(source, mirrorCircle.Center, mirrorCircle.Radius);
        }

        // The image of a point that is not there is not there either. (Asked only
        // whether its own coordinates were numbers - and the source keeps its last ones -
        // the image of an intersection that had gone stayed on screen, frozen, with all
        // that was built on it. Likewise a rotated, a dilated point, a point by
        // coordinates and an angle bisector.)
        Exists = Dependencies.Exists() && Coordinates.Exists();
    }
}