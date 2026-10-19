using System.Collections.Generic;
using System.Linq;
using GuiLabs.Undo;

namespace LiveGeometry;

public partial class Polygon : PolygonBase, IPolygon, ISupportRemoveDependency
{
    /// <summary>Triangle ABC, square ABCD; one with more vertices than that is numbered</summary>
    public const int MaxVerticesInName = 10;

    protected override IReadOnlyList<string> NamesFromDependencies()
    {
        return NamesFromPoints(PointOrder.Cyclic, MaxVerticesInName);
    }

    protected override string Kind
    {
        get
        {
            switch (Dependencies.Count)
            {
                case 3:
                    return "Triangle";
                case 4:
                    return Quadrilaterals.Classify(Point(0), Point(1), Point(2), Point(3));
                case 5:
                    return "Pentagon";
                case 6:
                    return "Hexagon";
                default:
                    return "Polygon";
            }
        }
    }

    // not offered (no [PropertyGridVisible], which the grid looks for on a method, whatever
    // it says): no tool makes a polyline (the Polyline tool is [Ignore]d), so one would be
    // a figure the user can't make or recognize
    [PropertyGridName("Convert to polyline")]
    [PropertyGridIcon(PropertyGridIcon.Polyline)]
    public void ConvertToPolyline()
    {
        // The outline through the same vertices and back to the first: ABCA. (It used to
        // be drawn on new free points, and the vertices deleted - with everything else
        // built on them, which in a construction is most of the drawing; redo of that threw.)
        var drawing = Drawing;
        var points = new List<IFigure>(Dependencies);
        points.Add(points[0]);
        var polyline = Factory.CreatePolyline(drawing, points);

        using (Transaction.Create(drawing.ActionManager, delayed: false))
        {
            Actions.Add(drawing, polyline);

            // in the polygon's place in the list
            Actions.MoveBefore(drawing, polyline, this);
            Actions.Remove(this);
        }

        drawing.RaiseUserIsAddingFigures(new Drawing.UIAFEventArgs() { Figures = polyline.AsEnumerable<IFigure>() });
    }

    #region Deleting a vertex

    /// <summary>
    /// A polygon loses the vertices deleted while three are left (ABCD without B is ACD), an
    /// image of one as well as the source: an image is a polygon on the images of the
    /// vertices, and loses the image of a vertex the source lost. Not a tool's preview, which
    /// goes with the point following the cursor as it always did (it would otherwise draw
    /// its closing side, outside the tool's figures); nor an arrow, which has no vertices.
    /// </summary>
    public bool CanRemoveDependencies(ISet<IFigure> removed)
    {
        var drawing = Drawing;
        if (drawing == null || drawing.ActionManager.ExecuteImmediatelyWithoutRecording || !Dependencies.All(vertex => vertex is IPoint))
        {
            return false;
        }

        if (Dependencies.Where(removed.Contains).Any(vertex => Dependencies.Count(other => other == vertex) != 1))
        {
            return false;
        }

        return Dependencies.Count(vertex => !removed.Contains(vertex)) >= 3;
    }

    public IAction GetRemoveDependencyAction(IFigure dependency)
    {
        return GetRemoveDependencyActions(new[] { dependency }).Single();
    }

    /// <summary>
    /// The vertices go, and where the polygon's outline was drawn by segments (the shape tools
    /// draw it so, and those along the gap go with the vertices) a segment of the same style
    /// closes the gap: ABCD with segments along its sides, B deleted, gets segment AC. One
    /// action, made once: redo brings back the same segment.
    /// </summary>
    public IEnumerable<IAction> GetRemoveDependencyActions(IList<IFigure> dependencies)
    {
        var drawing = Drawing;
        var gone = new HashSet<IFigure>(dependencies);
        var before = Dependencies.ToList();
        var after = before.Where(vertex => !gone.Contains(vertex)).ToList();
        var closingSides = new List<AddFigureAction>();
        int count = before.Count;
        for (int i = 0; i < count; i++)
        {
            // a vertex that stays followed by one that goes: the side to the next that stays
            if (gone.Contains(before[i]) || !gone.Contains(before[(i + 1) % count]))
            {
                continue;
            }

            int next = (i + 1) % count;
            while (gone.Contains(before[next]))
            {
                next = (next + 1) % count;
            }

            var from = (IPoint)before[i];
            var to = (IPoint)before[next];
            var drawn = Transformer.FindSideSegment(drawing, before[i], before[(i + 1) % count])
                ?? Transformer.FindSideSegment(drawing, before[(next - 1 + count) % count], before[next]);
            if (drawn == null || drawing.Figures.FindLine(from, to) != null)
            {
                continue;
            }

            var side = Factory.CreateSegment(drawing, from, to);
            side.Style = drawn.Style;
            closingSides.Add(new AddFigureAction(drawing, side));
        }

        var action = new CallMethodAction(
            () =>
            {
                SetVertices(after, gone, registered: false);
                closingSides.ForEach(side => side.Execute());
            },
            () =>
            {
                for (int i = closingSides.Count - 1; i >= 0; i--)
                {
                    closingSides[i].UnExecute();
                }

                SetVertices(before, gone, registered: true);
            });
        return new[] { action };
    }

    /// <summary>
    /// The polygon on the vertices, those that go listed with it or not, and worked out again
    /// with what is built on it. Without the Debug build's check of the list after it: other
    /// figures of the same deletion are half way (an image path whose source has let go of
    /// the handles its helper points are built on).
    /// </summary>
    void SetVertices(List<IFigure> vertices, HashSet<IFigure> going, bool registered)
    {
        Dependencies = vertices;
        foreach (var vertex in going)
        {
            if (registered)
            {
                vertex.Dependents.Add(this);
            }
            else
            {
                vertex.Dependents.Remove(this);
            }
        }

        this.RecalculateAndUpdateVisual();
        var dependents = DependencyAlgorithms.FindDescendants(figure => figure.Dependents, new IFigure[] { this });
        dependents.Reverse();
        foreach (var dependent in dependents)
        {
            dependent.RecalculateAndUpdateVisual();
        }
    }

    #endregion
}
