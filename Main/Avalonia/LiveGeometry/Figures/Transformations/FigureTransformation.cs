using System;
using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>
/// The transformation a figure is the image of, read off the points it is built on: every
/// one of them a rotated, translated, dilated or reflected point, all by the same sources
/// (one center and one angle, as the vertices of one rotated triangle share them), or the
/// traced image <see cref="Transformer"/> draws as a locus. The image is a plain figure
/// (a polygon on rotated points), so this is how its page and its line in the Figure List
/// say "rotation of ABC about O by 45°" and offer the angle (<see cref="Values"/>), and how
/// the Number holding the angle says whose it is (<see cref="TiedValues.RoleOf"/>).
/// </summary>
public class FigureTransformation
{
    FigureTransformation(IReadOnlyList<IFigure> points, IReadOnlyList<IFigure> sources)
    {
        Points = points;
        Sources = sources;
    }

    /// <summary>The image points the figure is built on, in its order</summary>
    public IReadOnlyList<IFigure> Points { get; }

    /// <summary>What each of them is the image of</summary>
    public IReadOnlyList<IFigure> Sources { get; }

    /// <summary>The figure transformed, when the drawing has one on the sources (the triangle ABC the image A'B'C' is of); null otherwise</summary>
    public IFigure Source { get; private set; }

    /// <summary>
    /// The point whose tied values are the transformation's (the angle, the distance and
    /// direction, the factor): any of the points, since they share their sources, so that
    /// an edit through it moves the whole image. Null for a reflection, which has none.
    /// </summary>
    public ITiedValues Values => Points[0] as ITiedValues;

    /// <summary>
    /// The transformation the figure is the image of; null for a figure that is none: a
    /// point (its own page says it), a figure not built on points of one transformation,
    /// a locus the user drew. Only the kinds a transformation makes (the lines, ellipses,
    /// polygons and paths it transforms through their points, the locus it traces).
    /// </summary>
    public static FigureTransformation Of(IFigure figure)
    {
        if (figure == null || figure.Drawing == null || figure is IPoint)
        {
            return null;
        }

        if (figure is Locus locus)
        {
            return OfTracedImage(locus);
        }

        if (!(figure is ILine || figure is IEllipse || figure is IPolygonalChain || figure is BezierPath))
        {
            return null;
        }

        var result = OfPoints(figure.Dependencies.Where(dependency => dependency is IPoint).ToList());
        if (result != null)
        {
            result.Source = FindFigureOn(result.Sources, figure.GetType()) ?? result.Source;
        }

        return result;
    }

    /// <summary>The kinds a transformation makes an image of through their points</summary>
    static bool IsTransformableKind(IFigure figure)
    {
        return figure is ILine || figure is IEllipse || figure is IPolygonalChain || figure is BezierPath;
    }

    /// <summary>
    /// A locus made by <see cref="Transformer"/> for a figure it can't transform through
    /// points: a hidden point slides on the source and the locus is the path of its image
    /// </summary>
    static FigureTransformation OfTracedImage(Locus locus)
    {
        if (locus.Dependencies.Count < 2
            || !(locus.Dependencies[1] is PointOnFigure sliding)
            || !sliding.Auxiliary
            || sliding.Dependencies.Count == 0)
        {
            return null;
        }

        var result = OfPoints(new[] { locus.Dependencies[0] });
        if (result == null || result.Sources[0] != sliding)
        {
            return null;
        }

        result.Source = sliding.Dependencies[0];
        return result;
    }

    /// <summary>
    /// The transformation the points are images of, when they all are, by the same sources;
    /// null otherwise. Reflected points go by their mirror; the others by what their
    /// <c>Siblings</c> share.
    /// </summary>
    public static FigureTransformation OfPoints(IReadOnlyList<IFigure> points)
    {
        if (points.Count == 0)
        {
            return null;
        }

        bool same;
        switch (points[0])
        {
            case RotatedPoint rotated:
                same = points.All(point => point is RotatedPoint other && other.Center == rotated.Center && other.AngleSource == rotated.AngleSource);
                break;
            case TranslatedPoint translated:
                same = translated.DistanceSource != null
                    && translated.DirectionSource != null
                    && points.All(point => point is TranslatedPoint other
                        && other.DistanceSource == translated.DistanceSource
                        && other.DirectionSource == translated.DirectionSource);
                break;
            case DilatedPoint dilated:
                same = points.All(point => point is DilatedPoint other
                    && other.Center == dilated.Center
                    && other.Dependencies.Skip(2).SequenceEqual(dilated.Dependencies.Skip(2)));
                break;
            case ReflectedPoint reflected:
                same = points.All(point => point is ReflectedPoint other && other.Mirror == reflected.Mirror);
                break;
            default:
                return null;
        }

        if (!same || points.Any(point => point.Dependencies.Count == 0))
        {
            return null;
        }

        var sources = points.Select(point => point.Dependencies[0]).ToList();
        return new FigureTransformation(points, sources) { Source = FindFigureOn(sources, kind: null) };
    }

    /// <summary>
    /// A figure on exactly the points, in their order - of the kind given, else of any kind
    /// a transformation makes an image of - else one of the kind given built on no other
    /// points (a path, whose handles the image has as points but the source may have as
    /// offsets); null when the drawing has none
    /// </summary>
    static IFigure FindFigureOn(IReadOnlyList<IFigure> sources, Type kind)
    {
        var candidates = sources[0].Dependents
            .Where(dependent => kind != null ? dependent.GetType() == kind : IsTransformableKind(dependent))
            .Distinct()
            .Select(dependent => (Figure: dependent, Points: dependent.Dependencies.Where(d => d is IPoint).ToList()))
            .ToList();
        var exact = candidates.FirstOrDefault(candidate => candidate.Points.SequenceEqual(sources)).Figure;
        if (exact != null || kind == null)
        {
            return exact;
        }

        var set = new HashSet<IFigure>(sources);
        return candidates.FirstOrDefault(candidate => candidate.Points.Count > 0 && candidate.Points.All(set.Contains)).Figure;
    }

    /// <summary>"rotation of ABC about O by 45°", "translation of AB by 3 at 30°", "dilation of c from O by 2", "reflection of ABC across line g", "inversion of line g in circle k"</summary>
    public string Describe()
    {
        switch (Points[0])
        {
            case RotatedPoint rotated:
                return "rotation" + DescribeSubject() + " by " + ConstructionText.AngleValue(rotated.AngleSource);
            case TranslatedPoint translated:
                return "translation" + DescribeSubject() + translated.MotionText;
            case DilatedPoint dilated:
                return "dilation" + DescribeSubject() + " by " + dilated.FactorText;
            case ReflectedPoint reflected:
                return (reflected.Mirror is ICircle ? "inversion" : "reflection") + DescribeSubject() + reflected.MirrorText;
            default:
                return null;
        }
    }

    /// <summary>
    /// What is transformed and about what: " of ABC about O" (a rotation), " of ABC from O"
    /// (a dilation), " of ABC" (a translation, a reflection); the points themselves where
    /// the drawing has no figure on them (" of A, B about O"), nothing of them when they
    /// are many or unnamed
    /// </summary>
    public string DescribeSubject()
    {
        string of = Source != null ? " of " + ConstructionText.Of(Source) : SourceNames();
        switch (Points[0])
        {
            case RotatedPoint rotated:
                return of + " about " + ConstructionText.Of(rotated.Center);
            case DilatedPoint dilated:
                return of + " from " + ConstructionText.Of(dilated.Center);
            default:
                return of;
        }
    }

    string SourceNames()
    {
        if (Sources.Count > 3 || Sources.Any(source => string.IsNullOrEmpty(source.Name)))
        {
            return "";
        }

        return " of " + string.Join(", ", Sources.Select(ConstructionText.Of));
    }

    /// <summary>The figure's construction line when it is an image (<see cref="Describe"/>), else null</summary>
    public static string Describe(IFigure figure)
    {
        return Of(figure)?.Describe();
    }
}
