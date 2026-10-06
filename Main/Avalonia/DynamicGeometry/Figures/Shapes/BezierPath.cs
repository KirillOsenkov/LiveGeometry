using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Collections;
using Avalonia.Controls;
using Avalonia.Media;
using GuiLabs.Undo;
using AvaloniaPath = Avalonia.Controls.Shapes.Path;
using AvaloniaShape = Avalonia.Controls.Shapes.Shape;

namespace DynamicGeometry;

/// <summary>
/// A path of cubic Bézier pieces through anchor points, as a WPF or Avalonia path is made of
/// Bézier segments. Each anchor has a handle on either side, kept as an offset from it (it
/// moves with its anchor), and the piece from one anchor to the next is pulled towards the
/// first one's out handle and the second one's in handle: a handle on its anchor leaves that
/// end straight, both make the piece a segment. Closed (a piece from the last anchor back to
/// the first) and filled are two switches of their own; an open path fills as if closed by a
/// straight line. A composite like a regular polygon: the anchors are points of the drawing
/// that it is built on, and its pieces (selected and styled one by one, the "sides"), its
/// inside and its handles are its parts. A handle shows only next to an anchor that is
/// selected or dragged (<see cref="IsHandleShown"/>), and only the Drag tool takes one:
/// nothing outside the path is built on a handle, but the images of a transformation.
/// Other paths can be its holes (<see cref="CutHoles"/>): the inside leaves them out, and
/// they stay paths of their own. The image of a transformation is a path whose handles are
/// points of the drawing (<see cref="CreateImage"/>), the images of the source's handles,
/// which follow those and can't be dragged themselves.
/// </summary>
public class BezierPath : CompositeFigure, IFigureParts, ILinearFigure, ISupportRemoveDependency
{
    readonly List<BezierPathHandle> inHandles = new List<BezierPathHandle>();
    readonly List<BezierPathHandle> outHandles = new List<BezierPathHandle>();
    readonly List<BezierPathPiece> pieces = new List<BezierPathPiece>();
    readonly Stack<BezierPathPiece> retiredPieces = new Stack<BezierPathPiece>();
    readonly BezierPathInterior interior;

    // the dotted lines from the anchors to the handles shown: a picture, not a figure
    readonly AvaloniaPath handleLines;

    // The dependencies are the anchors, then (for an image) two handle points per anchor,
    // in and out, then the holes. Whoever changes them keeps these counts right.
    int holeCount;
    bool handlePoints;

    bool closed;
    bool filled;

    // the anchor the Drag tool is dragging, whose handles show while it does
    IFigure draggedAnchor;

    public BezierPath()
    {
        interior = new BezierPathInterior(this);
        interior.Dependencies = new IFigure[] { this };
        Children.Add(interior);
        handleLines = new AvaloniaPath()
        {
            StrokeThickness = 1,
            StrokeDashArray = new AvaloniaList<double>() { 1, 2 },
            Opacity = HandleLineOpacity,
            IsHitTestVisible = false,
            IsVisible = false,
            ZIndex = (int)ZOrder.Figures + 1
        };
        handleLines.BindTheme(AvaloniaShape.StrokeProperty, nameof(AppTheme.Ink));
    }

    /// <summary>How much of the ink the lines from the anchors to their handles take</summary>
    public const double HandleLineOpacity = 0.3;

    /// <summary>
    /// A new path through the anchors, with handles at these offsets from them (zero: the
    /// handle is on its anchor), not in the drawing yet
    /// </summary>
    public static BezierPath Create(
        Drawing drawing,
        IList<IFigure> anchors,
        IList<Point> inOffsets,
        IList<Point> outOffsets,
        bool closed,
        bool filled)
    {
        var path = new BezierPath()
        {
            Drawing = drawing
        };
        path.closed = closed;
        path.filled = filled;
        path.Dependencies = anchors.ToList();
        path.SyncParts();
        for (int i = 0; i < anchors.Count; i++)
        {
            path.inHandles[i].Offset = inOffsets[i];
            path.outHandles[i].Offset = outOffsets[i];
        }

        return path;
    }

    #region Anchors, handles and pieces

    /// <summary>How many anchors the path goes through</summary>
    public int AnchorCount
    {
        get
        {
            return System.Math.Max(0, (Dependencies.Count - holeCount) / (handlePoints ? 3 : 1));
        }
    }

    /// <summary>How many pieces there are: one between each two anchors, and one back to the first when closed</summary>
    public int PieceCount
    {
        get
        {
            int count = AnchorCount;
            return count < 2 ? 0 : closed ? count : count - 1;
        }
    }

    public IPoint Anchor(int index)
    {
        return (IPoint)Dependencies[index];
    }

    /// <summary>Whether the figure is one of the anchors</summary>
    public bool IsAnchor(IFigure figure)
    {
        int count = AnchorCount;
        for (int i = 0; i < count; i++)
        {
            if (Dependencies[i] == figure)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>The paths whose insides the inside of this one leaves out</summary>
    public IEnumerable<BezierPath> Holes
    {
        get
        {
            return Dependencies.Skip(Dependencies.Count - holeCount).OfType<BezierPath>();
        }
    }

    /// <summary>Whether the handles are points of the drawing (an image's), which the Drag tool can't take</summary>
    public bool HasHandlePoints
    {
        get
        {
            return handlePoints;
        }
    }

    public IEnumerable<BezierPathHandle> Handles
    {
        get
        {
            return inHandles.Concat(outHandles);
        }
    }

    Point AnchorPoint(int index)
    {
        return ((IPoint)Dependencies[index]).Coordinates;
    }

    /// <summary>Where the handle on the side of the piece before the anchor is</summary>
    Point InPoint(int index)
    {
        return handlePoints
            ? ((IPoint)Dependencies[AnchorCount + 2 * index]).Coordinates
            : AnchorPoint(index).Plus(inHandles[index].Offset);
    }

    /// <summary>Where the handle on the side of the piece after the anchor is</summary>
    Point OutPoint(int index)
    {
        return handlePoints
            ? ((IPoint)Dependencies[AnchorCount + 2 * index + 1]).Coordinates
            : AnchorPoint(index).Plus(outHandles[index].Offset);
    }

    /// <summary>Where the handle is in the plane</summary>
    public Point HandleCoordinates(BezierPathHandle handle)
    {
        int index = handle.IsIn ? inHandles.IndexOf(handle) : outHandles.IndexOf(handle);
        if (index < 0 || index >= AnchorCount)
        {
            return handle.Coordinates;
        }

        return handle.IsIn ? InPoint(index) : OutPoint(index);
    }

    /// <summary>The handle across the anchor from this one</summary>
    BezierPathHandle Opposite(BezierPathHandle handle)
    {
        int index = handle.IsIn ? inHandles.IndexOf(handle) : outHandles.IndexOf(handle);
        if (index < 0)
        {
            return null;
        }

        return handle.IsIn ? outHandles[index] : inHandles[index];
    }

    /// <summary>The anchor the handle belongs to</summary>
    IFigure AnchorOf(BezierPathHandle handle)
    {
        int index = handle.IsIn ? inHandles.IndexOf(handle) : outHandles.IndexOf(handle);
        return index >= 0 && index < AnchorCount ? Dependencies[index] : null;
    }

    /// <summary>The handles of an anchor, as the tool that makes a path sets them while the path is a preview</summary>
    public void SetHandleOffsets(int index, Point inOffset, Point outOffset)
    {
        if (handlePoints || index < 0 || index >= inHandles.Count)
        {
            return;
        }

        inHandles[index].Offset = inOffset;
        outHandles[index].Offset = outOffset;
        RecalculateAndUpdate();
    }

    /// <summary>
    /// The handle goes where it is dragged: its offset from its anchor changes, and with
    /// <see cref="BezierPathHandle.MirrorsOpposite"/> the handle across the anchor is its
    /// mirror image through the anchor from then on (a smooth, symmetric anchor)
    /// </summary>
    public void MoveHandle(BezierPathHandle handle, Point coordinates)
    {
        var anchor = AnchorOf(handle);
        if (!(anchor is IPoint point))
        {
            return;
        }

        var offset = coordinates.Minus(point.Coordinates);
        if (handle.MirrorsOpposite && Opposite(handle) is BezierPathHandle opposite)
        {
            opposite.Offset = offset.Minus();
        }

        handle.Offset = offset;
        RecalculateAndUpdate();
    }

    /// <summary>For undo of a drag of a handle: where it and the one across the anchor are</summary>
    public object CaptureHandles(BezierPathHandle handle)
    {
        var opposite = Opposite(handle);
        return new Point[] { handle.Offset, opposite != null ? opposite.Offset : default };
    }

    public void RestoreHandles(BezierPathHandle handle, object place)
    {
        if (!(place is Point[] offsets) || offsets.Length != 2)
        {
            return;
        }

        handle.Offset = offsets[0];
        var opposite = Opposite(handle);
        if (opposite != null)
        {
            opposite.Offset = offsets[1];
        }

        RecalculateAndUpdate();
    }

    /// <summary>The path worked out and drawn again, and what is built on it (the points on it, a path it is a hole of)</summary>
    void RecalculateAndUpdate()
    {
        if (Drawing == null)
        {
            return;
        }

        this.RecalculateAndUpdateVisual();
        this.RecalculateAllDependents();
    }

    /// <summary>
    /// The parts the anchors need, made or taken away at the end: two handles per anchor
    /// (none for an image, whose handles are points of the drawing) and the pieces. A piece
    /// that leaves is kept, and is the one that comes back (undo of Closed).
    /// </summary>
    void SyncParts()
    {
        int count = AnchorCount;
        int handles = handlePoints ? 0 : count;
        while (inHandles.Count < handles)
        {
            var inHandle = new BezierPathHandle(this, isIn: true);
            var outHandle = new BezierPathHandle(this, isIn: false);
            inHandles.Add(inHandle);
            outHandles.Add(outHandle);
            AddPart(inHandle, CommonStyle(Handles.Where(h => h != inHandle && h != outHandle)));
            AddPart(outHandle, CommonStyle(Handles.Where(h => h != outHandle)));
        }

        while (inHandles.Count > handles)
        {
            RemovePart(inHandles[inHandles.Count - 1]);
            RemovePart(outHandles[outHandles.Count - 1]);
            inHandles.RemoveLast();
            outHandles.RemoveLast();
        }

        SyncPieces();
    }

    void SyncPieces()
    {
        int count = PieceCount;
        while (pieces.Count < count)
        {
            bool isNew = retiredPieces.Count == 0;
            var piece = isNew ? new BezierPathPiece(this) : retiredPieces.Pop();
            var common = isNew ? CommonStyle(pieces) : null;
            pieces.Add(piece);
            AddPart(piece, common);
        }

        while (pieces.Count > count)
        {
            var piece = pieces[pieces.Count - 1];
            pieces.RemoveLast();
            RemovePart(piece);
            retiredPieces.Push(piece);
        }
    }

    /// <summary>The style most of these parts have; null when none has one yet</summary>
    static IFigureStyle CommonStyle(IEnumerable<IFigure> parts)
    {
        return parts
            .Where(part => part.Style != null)
            .GroupBy(part => part.Style)
            .OrderByDescending(group => group.Count())
            .FirstOrDefault()?
            .Key;
    }

    #endregion

    #region Parts

    // The parts are figures of the library that the drawing never sees (as a regular
    // polygon's): no names, found through the path, built on it so that what is built on
    // them (only the images of a transformation) follows it. They are listed with the path
    // only while it is in the drawing (see DependentPolygonBase.RegisterPart).
    bool partsUnregistered = true;

    /// <summary>Whether the parts' shapes are on a canvas: parts made before the path is on one get there with it</summary>
    bool IsOnCanvas { get; set; }

    void AddPart(IFigure part, IFigureStyle style)
    {
        part.Dependencies = new IFigure[] { this };
        part.Drawing = Drawing;
        part.Visible = Visible;
        part.Selected = part is BezierPathPiece && Selected;
        if (!partsUnregistered)
        {
            part.RegisterWithDependencies();
        }

        Children.Add(part);
        if (style != null)
        {
            part.Style = style;
        }

        if (IsOnCanvas)
        {
            part.OnAddingToCanvas(Drawing.Canvas);
        }
    }

    void RemovePart(IFigure part)
    {
        part.UnregisterFromDependencies();
        Children.Remove(part);
        if (IsOnCanvas)
        {
            part.OnRemovingFromCanvas(Drawing.Canvas);
        }
    }

    /// <summary>A part that comes back (undo): where it was among the parts, listed again</summary>
    void ReturnPart(IFigure part)
    {
        if (!partsUnregistered)
        {
            part.RegisterWithDependencies();
        }

        Children.Add(part);
        if (IsOnCanvas)
        {
            part.OnAddingToCanvas(Drawing.Canvas);
        }
    }

    public override void OnAddingToDrawing(Drawing drawing)
    {
        base.OnAddingToDrawing(drawing);
        if (partsUnregistered)
        {
            partsUnregistered = false;
            foreach (var part in Children)
            {
                part.RegisterWithDependencies();
            }
        }

        drawing.SelectionChanged -= Drawing_SelectionChanged;
        drawing.SelectionChanged += Drawing_SelectionChanged;
    }

    public override void OnRemovingFromDrawing(Drawing drawing)
    {
        base.OnRemovingFromDrawing(drawing);
        drawing.SelectionChanged -= Drawing_SelectionChanged;
        if (!partsUnregistered)
        {
            partsUnregistered = true;
            foreach (var part in Children)
            {
                part.UnregisterFromDependencies();
            }
        }
    }

    public override void OnAddingToCanvas(Canvas newContainer)
    {
        base.OnAddingToCanvas(newContainer);
        newContainer.Children.Add(handleLines);
        IsOnCanvas = true;
    }

    public override void OnRemovingFromCanvas(Canvas leavingContainer)
    {
        base.OnRemovingFromCanvas(leavingContainer);
        leavingContainer.Children.Remove(handleLines);
        IsOnCanvas = false;
    }

    const string InPart = "In";
    const string OutPart = "Out";
    const string PiecePart = "Piece";
    const string InteriorPart = "Interior";

    /// <summary>In1, Out1 (the handles of the first anchor), Piece1 (from the first anchor to the second), Interior</summary>
    public string GetPartName(IFigure part)
    {
        int index = inHandles.IndexOf(part as BezierPathHandle);
        if (index >= 0)
        {
            return InPart + (index + 1);
        }

        index = outHandles.IndexOf(part as BezierPathHandle);
        if (index >= 0)
        {
            return OutPart + (index + 1);
        }

        index = pieces.IndexOf(part as BezierPathPiece);
        if (index >= 0)
        {
            return PiecePart + (index + 1);
        }

        return part == interior ? InteriorPart : null;
    }

    public IFigure GetPart(string partName)
    {
        if (partName == InteriorPart)
        {
            return interior;
        }

        if (TryGetIndex(partName, InPart, out int index))
        {
            return index >= 1 && index <= inHandles.Count ? inHandles[index - 1] : null;
        }

        if (TryGetIndex(partName, OutPart, out index))
        {
            return index >= 1 && index <= outHandles.Count ? outHandles[index - 1] : null;
        }

        if (TryGetIndex(partName, PiecePart, out index))
        {
            return index >= 1 && index <= pieces.Count ? pieces[index - 1] : null;
        }

        return null;
    }

    static bool TryGetIndex(string partName, string kind, out int index)
    {
        index = 0;
        return partName != null
            && partName.StartsWith(kind, StringComparison.Ordinal)
            && int.TryParse(partName.Substring(kind.Length), NumberStyles.None, CultureInfo.InvariantCulture, out index);
    }

    /// <summary>The pieces: a click selects one by itself, to be styled; the inside selects the path</summary>
    public IEnumerable<IFigure> SelectableParts
    {
        get
        {
            return pieces;
        }
    }

    /// <summary>"Side 2 of Bezier path ABC", "Handle of B toward C": the title of a part's page, what Tab says of it</summary>
    public string DescribePart(IFigure part)
    {
        int index = pieces.IndexOf(part as BezierPathPiece);
        if (index >= 0)
        {
            return "Side " + (index + 1) + " of " + Reference;
        }

        if (part is BezierPathHandle handle && AnchorOf(handle) is IFigure anchor)
        {
            int count = AnchorCount;
            int anchorIndex = Dependencies.IndexOf(anchor);
            int toward = (anchorIndex + (handle.IsIn ? count - 1 : 1)) % count;
            return "Handle of " + ConstructionText.Of(anchor) + " toward " + ConstructionText.Of(Dependencies[toward]);
        }

        return Reference;
    }

    /// <summary>The rows of a side's page: its style</summary>
    public IEnumerable<IValueProvider> GetPartProperties(IFigure part)
    {
        yield return PropertyDiscoveryStrategy.CreateValueProvider(part, nameof(StyleDisplay));
    }

    public IEnumerable<IOperationDescription> GetPartMethods(IFigure part)
    {
        yield return MethodDescription.Create(typeof(FigureBase).GetMethod(nameof(EditStyleButton)));
        yield return MethodDescription.Create(typeof(FigureBase).GetMethod(nameof(CreateNewStyle)));
        yield return new DelegateOperation(
            "SelectWhole",
            "Select " + Reference,
            PropertyGridIcon.Polygon,
            SelectWhole);
    }

    /// <summary>The path instead of a side of it, in the selection and in the grid</summary>
    public void SelectWhole()
    {
        if (Drawing == null)
        {
            return;
        }

        Drawing.Figures.ClearSelection();
        Selected = true;
        Drawing.RaiseSelectionChanged(Drawing.GetSelectedFigures());
    }

    #endregion

    #region Selection and visibility

    // whether the whole path is selected: a side selected by itself leaves this off (the
    // side is what the grid shows), while selecting the whole selects every side and the
    // inside, for their halos. Never the handles: a selection has no halo there.
    bool selectedWhole;

    public override bool Selected
    {
        get
        {
            return selectedWhole;
        }
        set
        {
            selectedWhole = value;
            interior.Selected = value;
            foreach (var piece in pieces)
            {
                piece.Selected = value;
            }
        }
    }

    void Drawing_SelectionChanged(object sender, Drawing.SelectionChangedEventArgs e)
    {
        RefreshHandles();
    }

    /// <summary>
    /// The Drag tool shows the handles next to the point it drags, selected or not, while it
    /// drags it (null: the drag is over)
    /// </summary>
    public static void ShowHandlesWhileDragging(Drawing drawing, IFigure point)
    {
        if (drawing == null)
        {
            return;
        }

        foreach (var path in drawing.Figures.OfType<BezierPath>())
        {
            var anchor = point != null && path.IsAnchor(point) ? point : null;
            if (path.draggedAnchor != anchor)
            {
                path.draggedAnchor = anchor;
                path.RefreshHandles();
            }
        }
    }

    /// <summary>
    /// Whether the handle shows: next to an anchor that is selected or dragged - its own two,
    /// and those of its neighbors that face it, which bend the same pieces - and only where
    /// it bends a piece (not the outer handles of an open path's ends)
    /// </summary>
    public bool IsHandleShown(BezierPathHandle handle)
    {
        if (!Exists || !Visible || Drawing == null)
        {
            return false;
        }

        int count = AnchorCount;
        int index = handle.IsIn ? inHandles.IndexOf(handle) : outHandles.IndexOf(handle);
        if (index < 0 || index >= count || count < 2)
        {
            return false;
        }

        int other = handle.IsIn ? index - 1 : index + 1;
        if (closed)
        {
            other = (other + count) % count;
        }
        else if (other < 0 || other >= count)
        {
            return false;
        }

        return IsActive(Dependencies[index]) || IsActive(Dependencies[other]);
    }

    bool IsActive(IFigure anchor)
    {
        return anchor.Selected || anchor == draggedAnchor;
    }

    /// <summary>The handles shown or hidden, where they are, and the lines to them</summary>
    void RefreshHandles()
    {
        foreach (var handle in Handles)
        {
            handle.Refresh();
        }

        UpdateHandleLines();
    }

    void UpdateHandleLines()
    {
        var shown = Handles.Where(handle => handle.Visible && handle.Exists).ToList();
        if (shown.Count == 0 || Drawing == null || !Exists)
        {
            handleLines.IsVisible = false;
            return;
        }

        var geometry = new PathGeometry();
        foreach (var handle in shown)
        {
            if (!(AnchorOf(handle) is IPoint anchor))
            {
                continue;
            }

            var from = ToPhysical(anchor.Coordinates);
            var to = ToPhysical(handle.Coordinates);
            if (!from.Exists() || !to.Exists())
            {
                continue;
            }

            geometry.Figures.Add(new PathFigure()
            {
                StartPoint = from,
                IsClosed = false,
                IsFilled = false,
                Segments = new PathSegmentCollection() { new LineSegment() { Point = to } }
            });
        }

        handleLines.Data = geometry;
        handleLines.IsVisible = true;
    }

    #endregion

    #region Properties

    /// <summary>Whether a piece goes from the last anchor back to the first</summary>
    [PropertyGridVisible]
    [PropertyGridName("Closed")]
    public bool Closed
    {
        get
        {
            return closed;
        }
        set
        {
            if (closed == value)
            {
                return;
            }

            closed = value;
            SyncPieces();
            if (Drawing != null)
            {
                RecalculateAndUpdate();
            }

            RaisePropertyChanged(nameof(Closed));
        }
    }

    /// <summary>Whether the inside is filled: an open path as if a straight line closed it</summary>
    [PropertyGridVisible]
    [PropertyGridName("Filled")]
    public bool Filled
    {
        get
        {
            return filled;
        }
        set
        {
            if (filled == value)
            {
                return;
            }

            filled = value;
            if (Drawing != null)
            {
                RecalculateAndUpdate();
            }

            RaisePropertyChanged(nameof(Filled));
        }
    }

    /// <summary>
    /// The style of the inside, which is the path's own: a click inside selects the path, and
    /// this is what its grid edits as the fill. The sides and handles have styles of their own.
    /// </summary>
    public override IFigureStyle Style
    {
        get
        {
            return interior.Style;
        }
        set
        {
            interior.Style = value;
        }
    }

    [PropertyGridName("Fill")]
    public override IFigureStyle StyleDisplay
    {
        get
        {
            return base.StyleDisplay;
        }
        set
        {
            base.StyleDisplay = value;
        }
    }

    /// <summary>All the sides at once; one of them by itself is selected with a click on it</summary>
    [PropertyGridVisible]
    public PartStylesValue SideStyles
    {
        get
        {
            return pieces.Count > 0 ? new PartStylesValue(nameof(SideStyles), "Sides", pieces) : null;
        }
    }

    /// <summary>All the handles at once (they are not selected by themselves)</summary>
    [PropertyGridVisible]
    public PartStylesValue HandleStyles
    {
        get
        {
            return inHandles.Count > 0 ? new PartStylesValue(nameof(HandleStyles), "Handles", Handles.ToList()) : null;
        }
    }

    protected override string Kind
    {
        get
        {
            return "Bezier path";
        }
    }

    /// <summary>Named as polygons are, when no points name it: p, q...</summary>
    protected override string FirstLetter
    {
        get
        {
            return "p";
        }
    }

    /// <summary>After its anchors, as a polygon (closed) or a polyline (open) is</summary>
    protected override IReadOnlyList<string> NamesFromDependencies()
    {
        var anchors = Dependencies.Take(AnchorCount).ToList();
        return NamesFromPoints(anchors, closed ? PointOrder.Cyclic : PointOrder.Reversible, Polygon.MaxVerticesInName);
    }

    public override Point Center
    {
        get
        {
            int count = AnchorCount;
            if (count == 0)
            {
                return new Point();
            }

            double x = 0;
            double y = 0;
            for (int i = 0; i < count; i++)
            {
                var point = AnchorPoint(i);
                x += point.X;
                y += point.Y;
            }

            return new Point(x / count, y / count);
        }
    }

    /// <summary>The name, not the dump of the parts a composite gives</summary>
    public override string ToString()
    {
        return Name;
    }

    #endregion

    #region Working it out

    Math.BezierInfo[] curves = Array.Empty<Math.BezierInfo>();

    // the four points of each piece: its anchor, the two handles it is pulled towards, the next anchor
    Point[][] controls = Array.Empty<Point[]>();
    bool curvesKnown;

    /// <summary>
    /// The pieces worked out, also before the first <see cref="Recalculate"/>: a point on the
    /// path that is read from a file asks where it is while the file is still being read
    /// </summary>
    Math.BezierInfo[] Curves
    {
        get
        {
            if (!curvesKnown)
            {
                CalculateCurves();
            }

            return curves;
        }
    }

    void CalculateCurves()
    {
        int count = PieceCount;
        int anchors = AnchorCount;
        var result = new Math.BezierInfo[count];
        var points = new Point[count][];
        for (int i = 0; i < count; i++)
        {
            int next = (i + 1) % anchors;
            points[i] = new[] { AnchorPoint(i), OutPoint(i), InPoint(next), AnchorPoint(next) };
            result[i] = new Math.BezierInfo(points[i][0], points[i][1], points[i][2], points[i][3]);
        }

        curves = result;
        controls = points;
        curvesKnown = true;
    }

    /// <summary>Whether every point the pieces are drawn through is somewhere</summary>
    bool CurvesExist()
    {
        foreach (var curve in Curves)
        {
            if (curve.Points == null || !curve.Points.All(point => point.Exists()))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>The path exists while its anchors (and an image's handle points) do; a hole that doesn't is left out</summary>
    public override void UpdateExistence()
    {
        int count = Dependencies.Count - holeCount;
        bool exists = AnchorCount >= 2;
        for (int i = 0; i < count && exists; i++)
        {
            exists = Dependencies[i].Exists;
        }

        Exists = exists;
        foreach (var part in Children)
        {
            part.UpdateExistence();
        }
    }

    public override void Recalculate()
    {
        CalculateCurves();
        foreach (var handle in Handles)
        {
            handle.Recalculate();
        }
    }

    public override void UpdateVisual()
    {
        if (Drawing == null)
        {
            return;
        }

        if (Exists && Visible && CurvesExist())
        {
            for (int i = 0; i < pieces.Count && i < curves.Length; i++)
            {
                pieces[i].Shape.Data = PieceGeometry(i);
            }

            if (interior.Exists)
            {
                interior.Shape.Data = InteriorGeometry();
            }
        }

        RefreshHandles();
    }

    PathGeometry PieceGeometry(int index)
    {
        return new PathGeometry() { Figures = new PathFigureCollection() { PieceFigure(index) } };
    }

    PathFigure PieceFigure(int index)
    {
        return new PathFigure()
        {
            StartPoint = ToPhysical(controls[index][0]),
            IsClosed = false,
            IsFilled = false,
            Segments = new PathSegmentCollection() { ToSegment(controls[index]) }
        };
    }

    BezierSegment ToSegment(Point[] points)
    {
        return new BezierSegment()
        {
            Point1 = ToPhysical(points[1]),
            Point2 = ToPhysical(points[2]),
            Point3 = ToPhysical(points[3])
        };
    }

    /// <summary>The sides as they are drawn, in pixels (for a halo along them)</summary>
    public PathGeometry SidesGeometry()
    {
        var geometry = new PathGeometry();
        if (Drawing == null || !Exists || !CurvesExist())
        {
            return geometry;
        }

        for (int i = 0; i < controls.Length; i++)
        {
            geometry.Figures.Add(PieceFigure(i));
        }

        return geometry;
    }

    /// <summary>The outline as one figure, closed (with a straight line, when the path is open), in pixels</summary>
    PathGeometry OutlineGeometry()
    {
        if (!curvesKnown)
        {
            CalculateCurves();
        }

        var figure = new PathFigure()
        {
            StartPoint = ToPhysical(controls[0][0]),
            IsClosed = true,
            IsFilled = true,
            Segments = new PathSegmentCollection()
        };
        foreach (var points in controls)
        {
            figure.Segments.Add(ToSegment(points));
        }

        return new PathGeometry()
        {
            FillRule = FillRule.EvenOdd,
            Figures = new PathFigureCollection() { figure }
        };
    }

    /// <summary>The inside: the outline as a polygon fills it (even-odd), the holes taken out</summary>
    Geometry InteriorGeometry()
    {
        var outline = OutlineGeometry();
        var holes = Holes.Where(hole => hole.Exists && hole.PieceCount > 0 && hole.CurvesExist()).ToList();
        if (holes.Count == 0)
        {
            return outline;
        }

        var group = new GeometryGroup() { FillRule = FillRule.NonZero };
        foreach (var hole in holes)
        {
            group.Children.Add(hole.OutlineGeometry());
        }

        return new CombinedGeometry(GeometryCombineMode.Exclude, outline, group);
    }

    /// <summary>The outline as a polygon through the points the pieces are drawn through</summary>
    List<Point> OutlinePolygon()
    {
        var result = new List<Point>();
        foreach (var curve in Curves)
        {
            if (curve.Points != null)
            {
                result.AddRange(curve.Points);
            }
        }

        return result;
    }

    /// <summary>Whether the point is inside: inside the outline (even-odd) and in none of the holes</summary>
    bool IsInside(Point point)
    {
        if (PieceCount == 0 || !OutlinePolygon().IsPointInPolygon(point))
        {
            return false;
        }

        return !Holes.Any(hole => hole.Exists && hole.PieceCount > 0 && hole.OutlinePolygon().IsPointInPolygon(point));
    }

    /// <summary>The smallest box around the pieces as drawn</summary>
    public Rect Bounds
    {
        get
        {
            var points = OutlinePolygon().Where(point => point.Exists()).ToList();
            if (points.Count == 0)
            {
                return new Rect();
            }

            double left = points.Min(p => p.X);
            double right = points.Max(p => p.X);
            double bottom = points.Min(p => p.Y);
            double top = points.Max(p => p.Y);
            return new Rect(left, bottom, right - left, top - bottom);
        }
    }

    #endregion

    #region Hit testing

    /// <summary>
    /// A handle that shows (for the Drag tool alone: no other tool may build on one), else a
    /// side, else the inside if it is filled. By the numbers, shown or not, as a vector's is:
    /// a point on a hidden path still exists, and hit testing leaves hidden figures out itself.
    /// </summary>
    public override IFigure HitTest(Point point)
    {
        if (Drawing == null || AnchorCount < 2)
        {
            return null;
        }

        if (Drawing.Behavior is Dragger)
        {
            foreach (var handle in Handles)
            {
                if (handle.Visible && handle.Exists && handle.HitTest(point) != null)
                {
                    return handle;
                }
            }
        }

        var curves = Curves;
        for (int i = 0; i < curves.Length && i < pieces.Count; i++)
        {
            // (the corners of the polyline count too: on the outer side of a bend a click is
            // over neither piece of it, see Bezier.HitTest)
            double reach = ToLogical(pieces[i].Shape.StrokeThickness / 2 + Math.CursorTolerance);
            if (curves[i].Points != null && Math.IsPointOnPolygonalChain(curves[i].Points, point, reach, false))
            {
                return pieces[i];
            }
        }

        if (filled && interior.Exists && IsInside(point))
        {
            return interior;
        }

        return null;
    }

    #endregion

    #region A point on the path

    /// <summary>What a point put on the figure under the cursor goes on: a side of a path stands for the path</summary>
    public static IFigure PointHolder(IFigure figure)
    {
        return figure is BezierPathPiece piece ? piece.Owner : figure;
    }

    /// <summary>
    /// A point on the path is on a piece: the whole number of the parameter is which (from
    /// 0), the rest how far along it the cubic's own parameter goes. A point on the closing
    /// piece of a path that is opened doesn't exist until it is closed again.
    /// </summary>
    public Tuple<double, double> GetParameterDomain()
    {
        return Tuple.Create(0.0, (double)PieceCount);
    }

    public Point GetPointFromParameter(double parameter)
    {
        var curves = Curves;
        const double slack = 1e-9;
        if (curves.Length == 0 || double.IsNaN(parameter) || parameter < -slack || parameter > curves.Length + slack)
        {
            return new Point(double.NaN, double.NaN);
        }

        parameter = System.Math.Max(0, System.Math.Min(curves.Length, parameter));
        int index = System.Math.Min((int)System.Math.Floor(parameter), curves.Length - 1);
        return curves[index].GetPoint(parameter - index);
    }

    public double GetNearestParameterFromPoint(Point point)
    {
        var curves = Curves;
        double bestDistance = double.MaxValue;
        int bestIndex = 0;
        double bestT = 0;
        for (int i = 0; i < curves.Length; i++)
        {
            var points = curves[i].Points;
            if (points == null)
            {
                continue;
            }

            for (int j = 0; j + 1 < points.Length; j++)
            {
                var (distance, ratio) = DistanceToSegment(point, points[j], points[j + 1]);
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    bestIndex = i;
                    bestT = (j + ratio) / (points.Length - 1);
                }
            }
        }

        if (curves.Length == 0)
        {
            return 0;
        }

        // the polyline is a few hundredths of the curve off: the nearest place on the curve
        // itself, in the step around the one found
        var best = curves[bestIndex];
        double step = 1.0 / (Math.BezierInfo.NumberOfPoints - 1);
        double low = System.Math.Max(0, bestT - step);
        double high = System.Math.Min(1, bestT + step);
        for (int k = 0; k < 40; k++)
        {
            double third = (high - low) / 3;
            if (best.GetPoint(low + third).Distance(point) < best.GetPoint(high - third).Distance(point))
            {
                high -= third;
            }
            else
            {
                low += third;
            }
        }

        // and a few steps of Newton's method on the squared distance, which the search above
        // leaves a little off (it is flat at its minimum)
        double t = (low + high) / 2;
        var c = controls[bestIndex];
        for (int k = 0; k < 4; k++)
        {
            double u = 1 - t;
            var at = best.GetPoint(t);
            double dx = at.X - point.X;
            double dy = at.Y - point.Y;
            double d1x = 3 * u * u * (c[1].X - c[0].X) + 6 * u * t * (c[2].X - c[1].X) + 3 * t * t * (c[3].X - c[2].X);
            double d1y = 3 * u * u * (c[1].Y - c[0].Y) + 6 * u * t * (c[2].Y - c[1].Y) + 3 * t * t * (c[3].Y - c[2].Y);
            double d2x = 6 * u * (c[2].X - 2 * c[1].X + c[0].X) + 6 * t * (c[3].X - 2 * c[2].X + c[1].X);
            double d2y = 6 * u * (c[2].Y - 2 * c[1].Y + c[0].Y) + 6 * t * (c[3].Y - 2 * c[2].Y + c[1].Y);
            double slope = dx * d1x + dy * d1y;
            double curvature = d1x * d1x + d1y * d1y + dx * d2x + dy * d2y;
            if (curvature <= 0)
            {
                break;
            }

            t = System.Math.Max(0, System.Math.Min(1, t - slope / curvature));
        }

        return bestIndex + t;
    }

    static (double Distance, double Ratio) DistanceToSegment(Point point, Point start, Point end)
    {
        double dx = end.X - start.X;
        double dy = end.Y - start.Y;
        double lengthSquared = dx * dx + dy * dy;
        double ratio = lengthSquared == 0
            ? 0
            : System.Math.Max(0, System.Math.Min(1, ((point.X - start.X) * dx + (point.Y - start.Y) * dy) / lengthSquared));
        var nearest = new Point(start.X + ratio * dx, start.Y + ratio * dy);
        return (nearest.Distance(point), ratio);
    }

    #endregion

    #region Removing an anchor or a hole

    /// <summary>
    /// A path keeps going without an anchor that is deleted, as long as two are left; an
    /// image's anchors are tied to its handle points, and it goes with them. A hole that is
    /// deleted leaves the inside whole again.
    /// </summary>
    public bool CanRemoveDependency(IFigure dependency)
    {
        int index = Dependencies.IndexOf(dependency);
        if (index < 0 || Dependencies.Count(d => d == dependency) != 1)
        {
            return false;
        }

        if (index >= Dependencies.Count - holeCount)
        {
            return true;
        }

        return !handlePoints && index < AnchorCount && AnchorCount > 2;
    }

    public IAction GetRemoveDependencyAction(IFigure dependency)
    {
        int index = Dependencies.IndexOf(dependency);
        if (index >= Dependencies.Count - holeCount)
        {
            var hole = (BezierPath)dependency;
            return new CallMethodAction(() => RemoveHole(hole), () => AddHole(hole));
        }

        return CreateRemoveAnchorAction(dependency, joinTarget: null);
    }

    /// <summary>
    /// Whether the point can be dropped from the path into the target, an anchor next to it
    /// (an Alt-drag of one onto the other, <see cref="PointSnapping.Join"/>), with two
    /// anchors left
    /// </summary>
    public bool CanDropAnchorInto(IFigure point, IFigure target)
    {
        int count = AnchorCount;
        if (handlePoints || count <= 2)
        {
            return false;
        }

        int index = Dependencies.IndexOf(point);
        int targetIndex = Dependencies.IndexOf(target);
        if (index < 0 || index >= count || targetIndex < 0 || targetIndex >= count)
        {
            return false;
        }

        int apart = System.Math.Abs(index - targetIndex);
        return apart == 1 || closed && apart == count - 1;
    }

    /// <summary>
    /// The point leaves the path for the target next to it: the piece between them goes, and
    /// the target takes the point's handle on the far side, so that the piece that joins the
    /// target to the point's other neighbor is shaped as the point's was. One undo step.
    /// </summary>
    public IAction CreateDropAnchorAction(IFigure point, IFigure target)
    {
        return CreateRemoveAnchorAction(point, target);
    }

    IAction CreateRemoveAnchorAction(IFigure anchor, IFigure joinTarget)
    {
        int index = -1;
        int pieceIndex = -1;
        BezierPathHandle inHandle = null;
        BezierPathHandle outHandle = null;
        BezierPathPiece piece = null;
        BezierPathHandle heir = null;
        Point heirOffset = default;
        List<(PointOnFigure Point, double Parameter)> moved = null;

        return new CallMethodAction(
            () =>
            {
                int count = AnchorCount;
                index = Dependencies.IndexOf(anchor);
                inHandle = inHandles[index];
                outHandle = outHandles[index];
                heir = null;
                if (joinTarget != null)
                {
                    int target = Dependencies.IndexOf(joinTarget);
                    bool targetIsNext = target == (index + 1) % count;
                    heir = targetIsNext ? inHandles[target] : outHandles[target];
                    heirOffset = heir.Offset;
                    heir.Offset = targetIsNext ? inHandle.Offset : outHandle.Offset;
                }

                // the piece after the anchor goes (the one before, at the end of an open
                // path), and the one left joins its neighbors
                pieceIndex = !closed && index == count - 1 ? index - 1 : index;
                piece = pieces[pieceIndex];
                moved = MovePointsOff(pieceIndex);
                pieces.RemoveAt(pieceIndex);
                RemovePart(piece);
                inHandles.RemoveAt(index);
                outHandles.RemoveAt(index);
                RemovePart(inHandle);
                RemovePart(outHandle);
                this.RemoveDependencyCore(index, anchor);
                this.RecalculateAllDependents();
            },
            () =>
            {
                // the parts first: the anchor back in the list works the path out at once
                inHandles.Insert(index, inHandle);
                outHandles.Insert(index, outHandle);
                ReturnPart(inHandle);
                ReturnPart(outHandle);
                pieces.Insert(pieceIndex, piece);
                ReturnPart(piece);
                this.InsertDependencyCore(index, anchor);
                if (heir != null)
                {
                    heir.Offset = heirOffset;
                }

                foreach (var (point, parameter) in moved)
                {
                    point.Parameter = parameter;
                }

                RecalculateAndUpdate();
            });
    }

    /// <summary>
    /// The points on the path are on the same pieces when the piece at the index goes: those
    /// after it a piece back, those on it onto the piece that takes its place, as far along
    /// it. Returns where they were, for undo.
    /// </summary>
    List<(PointOnFigure Point, double Parameter)> MovePointsOff(int pieceIndex)
    {
        var result = new List<(PointOnFigure Point, double Parameter)>();
        foreach (var point in Dependents.OfType<PointOnFigure>().Where(p => p.Dependencies.FirstOrDefault() == this))
        {
            var parameter = point.Parameter;
            result.Add((point, parameter));
            if (parameter >= pieceIndex + 1)
            {
                point.Parameter = parameter - 1;
            }
            else if (parameter >= pieceIndex)
            {
                point.Parameter = System.Math.Max(0, pieceIndex - 1) + (parameter - pieceIndex);
            }
        }

        return result;
    }

    void AddHole(BezierPath hole)
    {
        holeCount++;
        this.InsertDependencyCore(Dependencies.Count, hole);
        RecalculateAndUpdate();
    }

    void RemoveHole(BezierPath hole)
    {
        int index = Dependencies.IndexOf(hole);
        if (index < 0)
        {
            return;
        }

        holeCount--;
        this.RemoveDependencyCore(index, hole);
        RecalculateAndUpdate();
    }

    #endregion

    #region Inserting an anchor

    /// <summary>The piece a parameter is on and how far along it (the cubic's own t)</summary>
    (int Piece, double T) SplitPlace(double parameter)
    {
        int count = PieceCount;
        int piece = System.Math.Max(0, System.Math.Min(count - 1, (int)System.Math.Floor(parameter)));
        return (piece, parameter - piece);
    }

    /// <summary>
    /// Whether the point can become an anchor of the path it is on, where it is: a point on
    /// a path with handles of its own (not an image), somewhere along a piece and not at
    /// its ends, and not one a locus is drawn from
    /// </summary>
    public static bool CanBecomeAnchor(IFigure point)
    {
        if (!(point is PointOnFigure onFigure)
            || !(onFigure.Dependencies.FirstOrDefault() is BezierPath path)
            || path.handlePoints
            || path.Drawing == null
            || !onFigure.Exists
            || PointSnapping.IsHeldByLocus(onFigure))
        {
            return false;
        }

        const double slack = 1e-6;
        var (_, t) = path.SplitPlace(onFigure.Parameter);
        return t > slack && t < 1 - slack;
    }

    /// <summary>
    /// The point on the path becomes an anchor of it there: a free point in its place (its
    /// name, label and dependents, <see cref="Actions.ReplacePoint"/>), put into the path
    /// between the ends of its piece, which is split there (de Casteljau's construction:
    /// the two pieces draw the same curve). The points on the path stay where they are.
    /// One undo step.
    /// </summary>
    public static void ConvertToAnchor(PointOnFigure point)
    {
        if (!CanBecomeAnchor(point))
        {
            return;
        }

        var path = (BezierPath)point.Dependencies[0];
        var drawing = point.Drawing;
        var (piece, t) = path.SplitPlace(point.Parameter);
        bool selected = point.Selected;
        FreePoint anchor;
        using (Transaction.Create(drawing.ActionManager, false))
        {
            anchor = Factory.CreateFreePoint(drawing, point.Coordinates);
            Actions.ReplacePoint(point, anchor);

            // before the path, which is built on it from now on
            Actions.MoveBefore(drawing, anchor, path);
            drawing.ActionManager.RecordAction(path.CreateInsertAnchorAction(anchor, piece, t));
        }

        if (selected)
        {
            point.Selected = false;
            anchor.Selected = true;
            drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        }
    }

    IAction CreateInsertAnchorAction(IFigure anchor, int piece, double t)
    {
        BezierPathHandle inHandle = null;
        BezierPathHandle outHandle = null;
        BezierPathPiece newPiece = null;
        Point oldOut = default;
        Point oldIn = default;
        List<(PointOnFigure Point, double Parameter)> moved = null;
        int index = piece + 1;

        return new CallMethodAction(
            () =>
            {
                int count = AnchorCount;
                int next = (piece + 1) % count;
                var p0 = AnchorPoint(piece);
                var p1 = OutPoint(piece);
                var p2 = InPoint(next);
                var p3 = AnchorPoint(next);
                var p01 = Lerp(p0, p1, t);
                var p12 = Lerp(p1, p2, t);
                var p23 = Lerp(p2, p3, t);
                var p012 = Lerp(p01, p12, t);
                var p123 = Lerp(p12, p23, t);
                var at = Lerp(p012, p123, t);

                oldOut = outHandles[piece].Offset;
                oldIn = inHandles[next].Offset;
                outHandles[piece].Offset = p01.Minus(p0);
                inHandles[next].Offset = p23.Minus(p3);
                moved = MovePointsAcross(piece, t);

                bool isNew = inHandle == null;
                if (isNew)
                {
                    inHandle = new BezierPathHandle(this, isIn: true);
                    outHandle = new BezierPathHandle(this, isIn: false);
                    newPiece = new BezierPathPiece(this);
                }

                inHandle.Offset = p012.Minus(at);
                outHandle.Offset = p123.Minus(at);
                inHandles.Insert(index, inHandle);
                outHandles.Insert(index, outHandle);
                pieces.Insert(piece + 1, newPiece);
                if (isNew)
                {
                    AddPart(inHandle, outHandles[piece].Style);
                    AddPart(outHandle, outHandles[piece].Style);
                    AddPart(newPiece, pieces[piece].Style);
                }
                else
                {
                    ReturnPart(inHandle);
                    ReturnPart(outHandle);
                    ReturnPart(newPiece);
                }

                this.InsertDependencyCore(index, anchor);
                RecalculateAndUpdate();
            },
            () =>
            {
                pieces.Remove(newPiece);
                inHandles.Remove(inHandle);
                outHandles.Remove(outHandle);
                RemovePart(newPiece);
                RemovePart(inHandle);
                RemovePart(outHandle);
                this.RemoveDependencyCore(index, anchor);
                int next = (piece + 1) % AnchorCount;
                outHandles[piece].Offset = oldOut;
                inHandles[next].Offset = oldIn;
                foreach (var (point, parameter) in moved)
                {
                    point.Parameter = parameter;
                }

                RecalculateAndUpdate();
            });
    }

    static Point Lerp(Point from, Point to, double t)
    {
        return new Point(from.X + (to.X - from.X) * t, from.Y + (to.Y - from.Y) * t);
    }

    /// <summary>
    /// The points on the path stay where they are when the piece is split at t: those on
    /// later pieces a piece on, those on it onto the half they are on (the halves are the
    /// same cubic, run over [0, t] and [t, 1]). Returns where they were, for undo.
    /// </summary>
    List<(PointOnFigure Point, double Parameter)> MovePointsAcross(int piece, double t)
    {
        var result = new List<(PointOnFigure Point, double Parameter)>();
        foreach (var point in Dependents.OfType<PointOnFigure>().Where(p => p.Dependencies.FirstOrDefault() == this))
        {
            var parameter = point.Parameter;
            result.Add((point, parameter));
            if (parameter >= piece + 1)
            {
                point.Parameter = parameter + 1;
            }
            else if (parameter >= piece)
            {
                double along = parameter - piece;
                point.Parameter = along < t
                    ? piece + along / t
                    : piece + 1 + (along - t) / (1 - t);
            }
        }

        return result;
    }

    #endregion

    #region Holes

    /// <summary>Whether the figures selected can be made one path with holes: two Bezier paths or more</summary>
    public static bool CanCutHoles(IEnumerable<IFigure> figures)
    {
        return FigureParts.Wholes(figures).OfType<BezierPath>().Count() >= 2;
    }

    /// <summary>
    /// The largest of the paths (by the box around it) leaves the others out of its inside:
    /// they are its holes, as the inside of the letter B leaves out two. They stay paths of
    /// their own, but without a fill (unfilled, undoably), and come before it in the list,
    /// since it is built on them. One that is built on the largest can't be a hole of it.
    /// One undo step.
    /// </summary>
    public static void CutHoles(Drawing drawing, IEnumerable<IFigure> figures)
    {
        var paths = FigureParts.Wholes(figures).OfType<BezierPath>().ToList();
        if (paths.Count < 2 || drawing == null)
        {
            return;
        }

        var primary = paths.OrderByDescending(path => path.Bounds.Width * path.Bounds.Height).First();
        var holes = paths
            .Where(path => path != primary && !path.DependsOn(primary) && !primary.Dependencies.Contains(path))
            .ToList();
        if (holes.Count == 0)
        {
            drawing.RaiseStatusNotification("These paths are holes of " + primary.Name + " already.");
            return;
        }

        using (Transaction.Create(drawing.ActionManager, false))
        {
            foreach (var hole in holes)
            {
                Actions.MoveBefore(drawing, hole, primary);
                drawing.ActionManager.RecordAction(new CallMethodAction(
                    () => primary.AddHole(hole),
                    () => primary.RemoveHole(hole)));
                if (hole.Filled)
                {
                    Actions.SetProperty(drawing.ActionManager, new PropertyValue(nameof(Filled), hole), false);
                }
            }

            if (!primary.Filled)
            {
                Actions.SetProperty(drawing.ActionManager, new PropertyValue(nameof(Filled), primary), true);
            }
        }
    }

    #endregion

    #region Transformations

    /// <summary>
    /// The image under a transformation, made by <paramref name="transform"/> point by point
    /// (it gives the figures it makes for one, the image last): the images of the anchors,
    /// of the handles (hidden, auxiliary: they go with the image), and of the holes, and a
    /// path through them whose handles are those points. Its handles follow the source's and
    /// can't be dragged themselves. The image last.
    /// </summary>
    public List<IFigure> CreateImage(Func<IFigure, List<IFigure>> transform)
    {
        var result = new List<IFigure>();
        IFigure Transform(IFigure source, bool helper)
        {
            var images = transform(source);
            if (helper)
            {
                foreach (var image in images)
                {
                    image.Visible = false;
                    image.Auxiliary = true;
                }
            }

            result.AddRange(images);
            return images.Last();
        }

        int count = AnchorCount;
        var dependencies = new List<IFigure>();
        for (int i = 0; i < count; i++)
        {
            dependencies.Add(Transform(Dependencies[i], helper: false));
        }

        for (int i = 0; i < count; i++)
        {
            dependencies.Add(Transform(HandleSource(i, isIn: true), helper: true));
            dependencies.Add(Transform(HandleSource(i, isIn: false), helper: true));
        }

        var holes = Holes.ToList();
        foreach (var hole in holes)
        {
            dependencies.Add(Transform(hole, helper: false));
        }

        var path = new BezierPath()
        {
            Drawing = Drawing
        };
        path.handlePoints = true;
        path.holeCount = holes.Count;
        path.closed = closed;
        path.filled = filled;
        path.Dependencies = dependencies;
        path.SyncParts();
        path.Visible = Visible;
        path.Style = Style;
        for (int i = 0; i < pieces.Count && i < path.pieces.Count; i++)
        {
            path.pieces[i].Style = pieces[i].Style;
        }

        result.Add(path);
        return result;
    }

    /// <summary>The handle as a point: the part, or an image's point</summary>
    IFigure HandleSource(int index, bool isIn)
    {
        if (handlePoints)
        {
            return Dependencies[AnchorCount + 2 * index + (isIn ? 0 : 1)];
        }

        return isIn ? inHandles[index] : outHandles[index];
    }

    /// <summary>Whether the transformations can take the path: what it is built on can all be transformed</summary>
    public bool CanBeTransformed(Func<IFigure, bool> canBeTransformed)
    {
        return Dependencies.All(canBeTransformed);
    }

    #endregion

    #region File

    /// <summary>
    /// The anchors are the dependencies, and their handles a path as Avalonia writes one,
    /// a piece from each anchor to the next, the closing one included whether or not the
    /// path is closed (it keeps its handles): "L" for a piece whose two handles are on their
    /// anchors, else "C" and the offsets of the first anchor's out handle and of the second
    /// one's in handle. An image says HandlePoints instead: its handles are dependencies, two
    /// per anchor after the anchors. The holes come last (Holes says how many).
    /// </summary>
    public override void WriteXml(XmlWriter writer)
    {
        base.WriteXml(writer);
        if (closed)
        {
            writer.WriteAttributeBool("Closed", true);
        }

        if (filled)
        {
            writer.WriteAttributeBool("Filled", true);
        }

        if (holeCount > 0)
        {
            writer.WriteAttributeString("Holes", holeCount.ToString(CultureInfo.InvariantCulture));
        }

        if (handlePoints)
        {
            writer.WriteAttributeBool("HandlePoints", true);
        }
        else
        {
            writer.WriteAttributeString("Path", HandlesText());
        }

        WritePartStyles(writer);
    }

    string HandlesText()
    {
        var text = new StringBuilder();
        int count = inHandles.Count;
        for (int i = 0; i < count; i++)
        {
            if (i > 0)
            {
                text.Append(' ');
            }

            var outOffset = outHandles[i].Offset;
            var inOffset = inHandles[(i + 1) % count].Offset;
            if (outOffset == default && inOffset == default)
            {
                text.Append('L');
            }
            else
            {
                text.Append("C ")
                    .Append(Number(outOffset.X)).Append(',').Append(Number(outOffset.Y)).Append(' ')
                    .Append(Number(inOffset.X)).Append(',').Append(Number(inOffset.Y));
            }
        }

        return text.ToString();
    }

    static string Number(double value)
    {
        // (plus zero: a negative zero would be written "-0")
        return (value + 0.0).ToStringInvariant();
    }

    /// <summary>The offsets <see cref="HandlesText"/> wrote; what can't be read leaves a handle on its anchor</summary>
    void ReadHandles(string text)
    {
        if (string.IsNullOrEmpty(text) || inHandles.Count == 0)
        {
            return;
        }

        var tokens = text.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
        int count = inHandles.Count;
        int piece = 0;
        for (int i = 0; i < tokens.Length && piece < count; piece++)
        {
            if (tokens[i] == "C" && i + 2 < tokens.Length)
            {
                if (TryParsePoint(tokens[i + 1], out var outOffset))
                {
                    outHandles[piece].Offset = outOffset;
                }

                if (TryParsePoint(tokens[i + 2], out var inOffset))
                {
                    inHandles[(piece + 1) % count].Offset = inOffset;
                }

                i += 3;
            }
            else
            {
                i++;
            }
        }
    }

    static bool TryParsePoint(string text, out Point point)
    {
        point = default;
        var parts = text.Split(',');
        if (parts.Length == 2
            && double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var x)
            && double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var y)
            && x.IsValidValue()
            && y.IsValidValue())
        {
            point = new Point(x, y);
            return true;
        }

        return false;
    }

    public override void ReadXml(XElement element)
    {
        closed = element.ReadBool("Closed", false);
        filled = element.ReadBool("Filled", false);
        handlePoints = element.ReadBool("HandlePoints", false);
        int.TryParse(element.ReadString("Holes"), NumberStyles.None, CultureInfo.InvariantCulture, out holeCount);

        // a file that doesn't add up (made by hand) is read as a path of anchors alone
        int listed = Dependencies.Count - holeCount;
        if (holeCount < 0 || listed < 0 || handlePoints && listed % 3 != 0)
        {
            holeCount = 0;
            handlePoints = false;
        }

        SyncParts();
        ReadHandles(element.ReadString("Path"));

        // visibility and the style of the inside, onto the parts made
        base.ReadXml(element);
        ReadPartStyles(element);
        curvesKnown = false;
    }

    const string SidesElement = "Sides";
    const string HandlesElement = "Handles";
    const string PartElement = "Part";

    /// <summary>
    /// The styles of the sides and handles, as a regular polygon writes its parts' (see
    /// DependentPolygonBase.WritePartStyles): the style most of them have unless it is the
    /// default, and a part with another by itself
    /// </summary>
    void WritePartStyles(XmlWriter writer)
    {
        WritePartStyles(writer, SidesElement, pieces);
        WritePartStyles(writer, HandlesElement, Handles);
    }

    void WritePartStyles(XmlWriter writer, string elementName, IEnumerable<IFigure> parts)
    {
        var styled = parts.Where(part => part.Style != null).ToList();
        if (styled.Count == 0 || Drawing == null)
        {
            return;
        }

        var common = CommonStyle(styled);
        if (common != Drawing.StyleManager.AssignDefaultStyle(styled[0]))
        {
            writer.WriteStartElement(elementName);
            writer.WriteAttributeString("Style", common.Name);
            writer.WriteEndElement();
        }

        foreach (var part in styled.Where(part => part.Style != common))
        {
            writer.WriteStartElement(PartElement);
            writer.WriteAttributeString("Name", GetPartName(part));
            writer.WriteAttributeString("Style", part.Style.Name);
            writer.WriteEndElement();
        }
    }

    void ReadPartStyles(XElement element)
    {
        var manager = Drawing?.StyleManager;
        if (manager == null)
        {
            return;
        }

        void Apply(IFigure part, string styleName)
        {
            var style = styleName != null ? manager[styleName] : null;
            if (part != null && style != null && style.GetType().SupportsFigureType(part.GetType()))
            {
                part.Style = style;
            }
        }

        foreach (var piece in pieces)
        {
            Apply(piece, (string)element.Element(SidesElement)?.Attribute("Style"));
        }

        foreach (var handle in Handles)
        {
            Apply(handle, (string)element.Element(HandlesElement)?.Attribute("Style"));
        }

        foreach (var part in element.Elements(PartElement))
        {
            Apply(GetPart((string)part.Attribute("Name")), (string)part.Attribute("Style"));
        }
    }

    #endregion

    #region The parts

    /// <summary>
    /// A handle of an anchor: an offset from it, which the Drag tool changes. Not selected by
    /// itself (a click on it leaves the selection as it is, so that it stays shown); styled
    /// with the others on the path's page.
    /// </summary>
    public class BezierPathHandle : PointBase, IFigurePart
    {
        public BezierPathHandle(BezierPath owner, bool isIn)
        {
            Owner = owner;
            IsIn = isIn;
        }

        public BezierPath Owner { get; }

        IFigureParts IFigurePart.Owner => Owner;

        /// <summary>Whether the handle is on the side of the piece that comes to its anchor (else of the one that leaves it)</summary>
        public bool IsIn { get; }

        /// <summary>Where the handle is from its anchor</summary>
        public Point Offset { get; set; }

        /// <summary>
        /// Set by the Drag tool unless Alt is held: the handle across the anchor follows as
        /// the mirror image of this one through the anchor
        /// </summary>
        public bool MirrorsOpposite { get; set; }

        public override void OnAddingToDrawing(Drawing drawing)
        {
        }

        public override void OnRemovingFromDrawing(Drawing drawing)
        {
        }

        // over the figures, under the points: an anchor its handle is on is taken first,
        // the handle with Tab (ClickChoice)
        protected override int DefaultZOrder()
        {
            return (int)ZOrder.Points - 1;
        }

        public override bool Visible
        {
            get
            {
                return base.Visible && Owner.IsHandleShown(this);
            }
            set
            {
                base.Visible = value;
            }
        }

        public override bool AllowMove()
        {
            return !Owner.Locked && Owner.Drawing != null;
        }

        public override void MoveToCore(Point newLocation)
        {
            Owner.MoveHandle(this, newLocation);
        }

        public override object CapturePlace()
        {
            return Owner.CaptureHandles(this);
        }

        public override void RestorePlace(object place)
        {
            Owner.RestoreHandles(this, place);
        }

        public override void Recalculate()
        {
            Coordinates = Owner.HandleCoordinates(this);
        }

        /// <summary>Shown or hidden as the anchors next to it are selected, and where it is</summary>
        public void Refresh()
        {
            UpdateShapeVisibility();
            if (IsShown)
            {
                UpdateVisual();
            }
        }

        public override string ToString()
        {
            return Owner.DescribePart(this);
        }
    }

    /// <summary>A piece of the path, from an anchor to the next: a side, selected and styled by itself</summary>
    public class BezierPathPiece : ShapeBase<AvaloniaPath>, IFigurePart, ICustomPropertyProvider, ICustomMethodProvider
    {
        public BezierPathPiece(BezierPath owner)
        {
            Owner = owner;
        }

        public BezierPath Owner { get; }

        IFigureParts IFigurePart.Owner => Owner;

        public override void OnAddingToDrawing(Drawing drawing)
        {
        }

        protected override AvaloniaPath CreateShape()
        {
            return new AvaloniaPath()
            {
                Stroke = new SolidColorBrush(Colors.Black),
                StrokeThickness = 1
            };
        }

        protected override int DefaultZOrder()
        {
            return (int)ZOrder.Figures;
        }

        // the path finds its parts (BezierPath.HitTest)
        public override IFigure HitTest(Point point)
        {
            return Owner.HitTest(point) == this ? this : null;
        }

        public IEnumerable<IValueProvider> GetProperties()
        {
            return Owner.GetPartProperties(this);
        }

        public IEnumerable<IOperationDescription> GetMethods()
        {
            return Owner.GetPartMethods(this);
        }

        public override string ToString()
        {
            return Owner.DescribePart(this);
        }
    }

    /// <summary>The filled inside: to the user, the path itself (a click inside selects the path)</summary>
    public class BezierPathInterior : ShapeBase<AvaloniaPath>, IFigurePart
    {
        public BezierPathInterior(BezierPath owner)
        {
            Owner = owner;
        }

        public BezierPath Owner { get; }

        IFigureParts IFigurePart.Owner => Owner;

        public override void OnAddingToDrawing(Drawing drawing)
        {
        }

        protected override AvaloniaPath CreateShape()
        {
            return new AvaloniaPath();
        }

        protected override int DefaultZOrder()
        {
            return (int)ZOrder.Polygons;
        }

        /// <summary>There while the path is filled</summary>
        public override void UpdateExistence()
        {
            Exists = Owner.Exists && Owner.Filled && Owner.PieceCount > 0;
        }

        public override IFigure HitTest(Point point)
        {
            return Owner.HitTest(point) == this ? this : null;
        }

        public override string ToString()
        {
            return Owner.Reference;
        }
    }

    #endregion
}
