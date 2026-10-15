using Avalonia;

namespace LiveGeometry;

/// <summary>
/// A figure with an inside a point can be kept in (<see cref="PointInRegion"/>): a polygon,
/// a circle or an ellipse, a sector or a circular segment, a closed or filled Bezier path.
/// Whether it is filled on screen doesn't matter: the inside is the figure's, the fill is
/// the style's.
/// </summary>
public interface IRegion : IFigure
{
    /// <summary>Whether there is an inside right now (an open, unfilled path has none)</summary>
    bool HasInside { get; }

    /// <summary>The area of the inside (which of two regions around a click is the smaller)</summary>
    double Area { get; }

    bool ContainsPoint(Point point);

    /// <summary>
    /// The point itself when it is inside, else the point of the outline nearest to it
    /// (a circle's along the ray from its center); a point that is nowhere when the figure
    /// has no inside to speak of.
    /// </summary>
    Point NearestPointInside(Point point);
}
