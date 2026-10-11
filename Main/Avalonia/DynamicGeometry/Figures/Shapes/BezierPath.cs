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
/// Bézier segments. Each anchor has a handle on either side, and the piece from one anchor to
/// the next is pulled towards the first one's out handle and the second one's in handle: a
/// handle on its anchor leaves that end straight, both make the piece a segment. A handle is
/// either an offset from its anchor (it moves with it, and the Drag tool drags it) or a point
/// of the drawing, which the path is built on (a handle on a tangent, the images of a
/// transformation's source handles). Closed (a piece from the last anchor back to the first)
/// and filled are two switches of their own; an open path fills as if closed by a straight
/// line. A composite like a regular polygon: the anchors are points of the drawing that it is
/// built on, and its pieces (selected and styled one by one, the "sides"), its inside and its
/// handles are its parts. A handle shows only next to an anchor that is selected or dragged
/// (<see cref="IsHandleShown"/>), and only the Drag tool takes one: nothing outside the path
/// is built on a handle, but the images of a transformation. Other paths can be its holes
/// (<see cref="CutHoles"/>): the inside leaves them out, and they stay paths of their own.
/// Whatever changes what the path is made of (an anchor, a hole, a handle's point) is one
/// change of its <see cref="Layout"/>, done and undone whole. A handle may be left to the path
/// (<see cref="BezierPathHandle.Auto"/>): it is worked out from the anchors by the path's
/// <see cref="Smoothing"/> (<see cref="BezierPathSmoother"/>), so the curve stays smooth as
/// they move; the tension it takes is typed or comes from a figure (a slider).
/// </summary>
public class BezierPath : CompositeFigure, IFigureParts, ILinearFigure, IPerimeter, ISupportRemoveDependency, ITiedValues, IConditionalProperties
{
    readonly List<BezierPathHandle> inHandles = new List<BezierPathHandle>();
    readonly List<BezierPathHandle> outHandles = new List<BezierPathHandle>();
    readonly List<BezierPathPiece> pieces = new List<BezierPathPiece>();
    readonly Stack<BezierPathPiece> retiredPieces = new Stack<BezierPathPiece>();
    readonly BezierPathInterior interior;

    // the dotted lines from the anchors to the handles shown: a picture, not a figure
    readonly AvaloniaPath handleLines;

    bool closed;
    bool filled;
    BezierPathSmoothing smoothing = BezierPathSmoothing.Hobby;

    // the tension typed, and the figure it comes from instead (the last dependency), if any
    double tension = DefaultTension;
    IFigure tensionSource;

    // the anchor the Drag tool is dragging, whose handles show while it does
    IFigure draggedAnchor;

    public BezierPath()
    {
        // the path itself is drawn by its parts; its layer says the verbs apply to it
        Layer = ZOrder.Polygons;
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
            ZIndex = ZOrders.Default(ZOrder.Handles) - 1
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
        return Create(
            drawing,
            anchors,
            anchors.Select((anchor, i) => new HandleSpec(inOffsets[i], null)).ToList(),
            anchors.Select((anchor, i) => new HandleSpec(outOffsets[i], null)).ToList(),
            holes: Array.Empty<IFigure>(),
            closed,
            filled);
    }

    /// <summary>
    /// What a handle is: an offset from its anchor, a point of the drawing (then the offset
    /// means nothing), or left to the path (automatic; the offset is what it was worked out to)
    /// </summary>
    public record struct HandleSpec(Point Offset, IPoint Point, bool Auto = false)
    {
        /// <summary>A handle the path works out</summary>
        public static HandleSpec Automatic => new HandleSpec(default, null, Auto: true);
    }

    /// <summary>A new path through the anchors, with these handles and holes, not in the drawing yet</summary>
    public static BezierPath Create(
        Drawing drawing,
        IList<IFigure> anchors,
        IList<HandleSpec> ins,
        IList<HandleSpec> outs,
        IList<IFigure> holes,
        bool closed,
        bool filled)
    {
        var path = new BezierPath()
        {
            Drawing = drawing
        };
        path.closed = closed;
        path.filled = filled;
        path.Build(anchors, ins, outs, holes);
        return path;
    }

    /// <summary>The handles, the pieces and the dependencies, for a path made or read</summary>
    void Build(IList<IFigure> anchors, IList<HandleSpec> ins, IList<HandleSpec> outs, IList<IFigure> holes)
    {
        for (int i = 0; i < anchors.Count; i++)
        {
            var inHandle = new BezierPathHandle(this, isIn: true);
            var outHandle = new BezierPathHandle(this, isIn: false);
            inHandle.Set(ins[i]);
            outHandle.Set(outs[i]);
            inHandles.Add(inHandle);
            outHandles.Add(outHandle);
            AttachPart(inHandle);
            AttachPart(outHandle);
        }

        SyncPieces();
        SetDependencies(DependenciesOf(anchors, holes, tensionSource));
    }

    /// <summary>The anchors, the points of the handles (<see cref="HandlesInOrder"/>), the holes, the figure the tension comes from</summary>
    List<IFigure> DependenciesOf(IEnumerable<IFigure> anchors, IEnumerable<IFigure> holes, IFigure source)
    {
        var result = anchors
            .Concat(HandlesInOrder.Where(h => h.Point != null).Select(h => (IFigure)h.Point))
            .Concat(holes)
            .ToList();
        if (source != null)
        {
            result.Add(source);
        }

        return result;
    }

    #region Anchors, handles and pieces

    /// <summary>How many anchors the path goes through</summary>
    public int AnchorCount
    {
        get
        {
            return inHandles.Count;
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
        int count = System.Math.Min(AnchorCount, Dependencies.Count);
        for (int i = 0; i < count; i++)
        {
            if (Dependencies[i] == figure)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// The handles in the order their points come among the dependencies, after the
    /// anchors: the in and the out handle of the first anchor, of the second...
    /// </summary>
    IEnumerable<BezierPathHandle> HandlesInOrder
    {
        get
        {
            for (int i = 0; i < inHandles.Count; i++)
            {
                yield return inHandles[i];
                yield return outHandles[i];
            }
        }
    }

    /// <summary>Where the holes start among the dependencies: after the anchors and the points of handles</summary>
    int HoleStart
    {
        get
        {
            return AnchorCount + HandlesInOrder.Count(h => h.Point != null);
        }
    }

    /// <summary>The dependencies after the points of handles, but for the figure the tension comes from (the last)</summary>
    List<IFigure> HoleList
    {
        get
        {
            int start = HoleStart;
            return Dependencies.Skip(start).Take(Dependencies.Count - start - (tensionSource != null ? 1 : 0)).ToList();
        }
    }

    /// <summary>The paths whose insides the inside of this one leaves out</summary>
    public IEnumerable<BezierPath> Holes
    {
        get
        {
            return HoleList.OfType<BezierPath>();
        }
    }

    /// <summary>Whether a handle is a point of the drawing</summary>
    public bool HasHandlePoints
    {
        get
        {
            return Handles.Any(h => h.Point != null);
        }
    }

    /// <summary>
    /// Whether the path is the image of a transformation: its handles are hidden helper
    /// points, which go with it. It isn't split or made shorter (the helpers would stay).
    /// </summary>
    bool IsImage
    {
        get
        {
            return Handles.Any(h => h.Point is IFigure point && point.Auxiliary);
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

    int IndexOf(BezierPathHandle handle)
    {
        return handle.IsIn ? inHandles.IndexOf(handle) : outHandles.IndexOf(handle);
    }

    /// <summary>Where the handle is in the plane: its point, or its offset from its anchor</summary>
    Point HandlePosition(BezierPathHandle handle, int index)
    {
        return handle.Point != null ? handle.Point.Coordinates : AnchorPoint(index).Plus(handle.Offset);
    }

    /// <summary>Where the handle on the side of the piece before the anchor is</summary>
    Point InPoint(int index)
    {
        return HandlePosition(inHandles[index], index);
    }

    /// <summary>Where the handle on the side of the piece after the anchor is</summary>
    Point OutPoint(int index)
    {
        return HandlePosition(outHandles[index], index);
    }

    /// <summary>Where the handle is in the plane</summary>
    public Point HandleCoordinates(BezierPathHandle handle)
    {
        int index = IndexOf(handle);
        if (index < 0 || index >= AnchorCount || Dependencies.Count < AnchorCount)
        {
            return handle.Coordinates;
        }

        return HandlePosition(handle, index);
    }

    /// <summary>The handle across the anchor from this one</summary>
    BezierPathHandle Opposite(BezierPathHandle handle)
    {
        int index = IndexOf(handle);
        if (index < 0)
        {
            return null;
        }

        return handle.IsIn ? outHandles[index] : inHandles[index];
    }

    /// <summary>The anchor the handle belongs to</summary>
    IFigure AnchorOf(BezierPathHandle handle)
    {
        int index = IndexOf(handle);
        return index >= 0 && index < AnchorCount && index < Dependencies.Count ? Dependencies[index] : null;
    }

    /// <summary>The handles of an anchor, as the tool that makes a path sets them while the path is a preview</summary>
    public void SetHandleOffsets(int index, Point inOffset, Point outOffset)
    {
        if (index < 0 || index >= inHandles.Count)
        {
            return;
        }

        inHandles[index].Set(new HandleSpec(inOffset, null));
        outHandles[index].Set(new HandleSpec(outOffset, null));
        RecalculateAndUpdate();
    }

    /// <summary>
    /// The handle goes where it is dragged: its offset from its anchor changes (it is the
    /// user's from then on, not automatic), and with <see cref="BezierPathHandle.MirrorsOpposite"/>
    /// the handle across the anchor is its mirror image through the anchor from then on (a
    /// smooth, symmetric anchor) - unless that one is a point of the drawing, which stays where
    /// it is. Without it an automatic one across stays where it was worked out to (a corner).
    /// </summary>
    public void MoveHandle(BezierPathHandle handle, Point coordinates)
    {
        var anchor = AnchorOf(handle);
        if (!(anchor is IPoint point) || handle.Point != null)
        {
            return;
        }

        var offset = coordinates.Minus(point.Coordinates);
        if (Opposite(handle) is BezierPathHandle { Point: null } opposite)
        {
            if (handle.MirrorsOpposite)
            {
                opposite.Offset = offset.Minus();
            }

            opposite.Auto = false;
        }

        handle.Offset = offset;
        handle.Auto = false;
        RecalculateAndUpdate();
    }

    /// <summary>For undo of a drag of a handle: what it and the one across the anchor are</summary>
    public object CaptureHandles(BezierPathHandle handle)
    {
        var opposite = Opposite(handle);
        return new HandleSpec[] { handle.Spec, opposite != null ? opposite.Spec : default };
    }

    public void RestoreHandles(BezierPathHandle handle, object place)
    {
        if (!(place is HandleSpec[] specs) || specs.Length != 2)
        {
            return;
        }

        handle.Set(specs[0]);
        var opposite = Opposite(handle);
        if (opposite != null)
        {
            opposite.Set(specs[1]);
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
    /// The same, without the Debug build's check of the list after it: an anchor deleted
    /// takes its handles off the path (<see cref="RemoveFigureAction"/> runs it first) while
    /// the helper points of an image built on them are still in the list, to go in the same
    /// deletion with the image of the anchor
    /// </summary>
    void RecalculateAndUpdateUnchecked()
    {
        if (Drawing == null)
        {
            return;
        }

        this.RecalculateAndUpdateVisual();
        var dependents = DependencyAlgorithms.FindDescendants(f => f.Dependents, new IFigure[] { this });
        dependents.Reverse();
        foreach (var dependent in dependents)
        {
            dependent.RecalculateAndUpdateVisual();
        }
    }

    /// <summary>
    /// What took the place of a point that is a handle (it was joined into another point,
    /// let go, replaced) is that handle's point from now on: the dependencies say, in the
    /// order of <see cref="HandlesInOrder"/>
    /// </summary>
    protected override void OnDependenciesChanged()
    {
        base.OnDependenciesChanged();
        int index = AnchorCount;
        foreach (var handle in HandlesInOrder.Where(h => h.Point != null))
        {
            if (index < Dependencies.Count && Dependencies[index] is IPoint point)
            {
                handle.Point = point;
            }

            index++;
        }

        if (tensionSource != null && Dependencies.Count > 0)
        {
            tensionSource = Dependencies[Dependencies.Count - 1];
        }

        curvesKnown = false;
    }

    /// <summary>The dependencies set, and listed with what they are, if they were</summary>
    void SetDependencies(List<IFigure> dependencies)
    {
        bool registered = Dependencies.Count > 0 && Dependencies.All(d => d.Dependents.Contains(this));
        if (registered)
        {
            this.UnregisterFromDependencies();
        }

        Dependencies = dependencies;
        if (registered)
        {
            this.RegisterWithDependencies();
        }
    }

    /// <summary>
    /// The pieces the anchors need, made or taken away at the end. A piece that leaves is
    /// kept, and is the one that comes back (undo of Closed).
    /// </summary>
    void SyncPieces()
    {
        int count = PieceCount;
        while (pieces.Count < count)
        {
            bool isNew = retiredPieces.Count == 0;
            var piece = isNew ? new BezierPathPiece(this) : retiredPieces.Pop();
            var common = isNew ? CommonStyle(pieces) : null;
            if (common != null)
            {
                piece.Style = common;
            }

            pieces.Add(piece);
            AttachPart(piece);
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

    #region Layout

    /// <summary>
    /// What the path is made of: its anchors, their handles (each an offset or a point),
    /// the pieces, the holes, and where the points on it are along it. A change is worked
    /// out on a copy (<see cref="LayoutChange"/>) and put in place whole.
    /// </summary>
    public class Layout
    {
        public List<IFigure> Anchors;
        public List<BezierPathHandle> Ins;
        public List<BezierPathHandle> Outs;
        public List<BezierPathPiece> Pieces;
        public Dictionary<BezierPathHandle, HandleSpec> Handles;
        public List<IFigure> Holes;
        public Dictionary<PointOnFigure, double> Parameters;
        public IFigure TensionSource;
        public double Tension;
    }

    Layout Capture()
    {
        return new Layout()
        {
            Anchors = Dependencies.Take(AnchorCount).ToList(),
            Ins = inHandles.ToList(),
            Outs = outHandles.ToList(),
            Pieces = pieces.ToList(),
            Handles = Handles.ToDictionary(h => h, h => h.Spec),
            Holes = HoleList,
            Parameters = PointsOnPath().ToDictionary(p => p, p => p.Parameter),
            TensionSource = tensionSource,
            Tension = tension
        };
    }

    IEnumerable<PointOnFigure> PointsOnPath()
    {
        return Dependents.OfType<PointOnFigure>().Where(p => p.Dependencies.FirstOrDefault() == this).Distinct();
    }

    /// <summary>The path made of what the layout says: parts that leave go off the canvas, new ones come on</summary>
    void Apply(Layout layout)
    {
        var parts = new HashSet<IFigure>(layout.Ins.Concat<IFigure>(layout.Outs).Concat(layout.Pieces));
        foreach (var part in Handles.Concat<IFigure>(pieces).Where(part => !parts.Contains(part)).ToList())
        {
            RemovePart(part);
        }

        var present = new HashSet<IFigure>(Children);
        inHandles.SetItems(layout.Ins);
        outHandles.SetItems(layout.Outs);
        pieces.SetItems(layout.Pieces);
        foreach (var pair in layout.Handles)
        {
            pair.Key.Set(pair.Value);
        }

        foreach (var part in layout.Ins.Concat<IFigure>(layout.Outs).Concat(layout.Pieces).Where(part => !present.Contains(part)))
        {
            AttachPart(part);
        }

        foreach (var pair in layout.Parameters)
        {
            pair.Key.Parameter = pair.Value;
        }

        tensionSource = layout.TensionSource;
        tension = layout.Tension;
        SetDependencies(DependenciesOf(layout.Anchors, layout.Holes, layout.TensionSource));
        RecalculateAndUpdateUnchecked();

        // the rows and buttons of the grid may change (a tension typed or tied, handles to smooth)
        RaisePropertyChanged(null);
    }

    /// <summary>
    /// A change of what the path is made of, as one action: worked out on a copy of the
    /// layout when it is first done (from the path as it is then), put in place, and the
    /// layout before put back on undo
    /// </summary>
    IAction LayoutChange(Action<Layout> change)
    {
        Layout before = null;
        Layout after = null;
        return new CallMethodAction(
            () =>
            {
                if (after == null)
                {
                    before = Capture();
                    after = Capture();
                    change(after);
                }

                Apply(after);
            },
            () => Apply(before));
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

    /// <summary>A part made or coming back: built on the path, listed with it, on its canvas</summary>
    void AttachPart(IFigure part)
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
            int anchorIndex = IndexOf(handle);
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
    /// drags it - or next to the anchor of a handle it drags (one taken with Tab while it
    /// didn't show). Null: the drag is over.
    /// </summary>
    public static void ShowHandlesWhileDragging(Drawing drawing, IFigure point)
    {
        if (drawing == null)
        {
            return;
        }

        foreach (var path in drawing.Figures.OfType<BezierPath>())
        {
            var dragged = point is BezierPathHandle handle && handle.Owner == path ? path.AnchorOf(handle) : point;
            var anchor = dragged != null && path.IsAnchor(dragged) ? dragged : null;
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
    /// it bends a piece (not the outer handles of an open path's ends). A handle that is a
    /// point of the drawing doesn't: the point shows itself.
    /// </summary>
    public bool IsHandleShown(BezierPathHandle handle)
    {
        return handle.Point == null && IsNextToActiveAnchor(handle);
    }

    bool IsNextToActiveAnchor(BezierPathHandle handle)
    {
        return BendsPiece(handle, out int index, out int other)
            && (IsActive(Dependencies[index]) || IsActive(Dependencies[other]));
    }

    /// <summary>
    /// Whether the handle bends a piece (not the outer handle of an open path's end), and
    /// the anchors of that piece: its own and the other one
    /// </summary>
    bool BendsPiece(BezierPathHandle handle, out int index, out int other)
    {
        int count = AnchorCount;
        index = IndexOf(handle);
        other = -1;
        if (!Exists || !Visible || Drawing == null || Dependencies.Count < count || index < 0 || count < 2)
        {
            return false;
        }

        other = handle.IsIn ? index - 1 : index + 1;
        if (closed)
        {
            other = (other + count) % count;
        }
        else if (other < 0 || other >= count)
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// The anchor's own handles that bend a piece and are at the point, shown or not: a
    /// handle on its anchor (a path of clicks) is there to be taken with Tab before the
    /// anchor is selected (<see cref="Dragger"/>)
    /// </summary>
    public IEnumerable<BezierPathHandle> HandlesOn(IFigure anchor, Point point)
    {
        int index = Dependencies.IndexOf(anchor);
        if (index < 0 || index >= AnchorCount || Drawing == null)
        {
            yield break;
        }

        var reach = Drawing.CoordinateSystem.CursorTolerance + ToLogical(HandleReach);
        foreach (var handle in new[] { inHandles[index], outHandles[index] })
        {
            if (handle.Point == null
                && handle.Exists
                && BendsPiece(handle, out _, out _)
                && HandleCoordinates(handle).Distance(point) <= reach)
            {
                yield return handle;
            }
        }
    }

    /// <summary>In pixels, besides the cursor's tolerance: how far from the cursor a handle hidden under its anchor is taken to be there</summary>
    const double HandleReach = 4;

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

    /// <summary>A dotted line from an anchor to each handle shown next to it, and to each point that is a handle there</summary>
    void UpdateHandleLines()
    {
        var shown = Handles
            .Where(handle => handle.Point == null
                ? handle.Visible && handle.Exists
                : handle.Point.Visible && handle.Point.Exists && IsNextToActiveAnchor(handle))
            .ToList();
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
            var to = ToPhysical(HandleCoordinates(handle));
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

    /// <summary>
    /// The length around a closed path, its pieces added up (<see cref="Math.BezierInfo.Length"/>);
    /// NaN for an open one, which has nothing to measure (the holes are paths of their own)
    /// </summary>
    public double Perimeter
    {
        get
        {
            if (!closed || PieceCount == 0 || !CurvesExist())
            {
                return double.NaN;
            }

            double sum = 0;
            foreach (var curve in Curves)
            {
                sum += curve.Length;
            }

            return sum;
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

    /// <summary>How the automatic handles are worked out (<see cref="BezierPathSmoother"/>); None leaves them on their anchors</summary>
    [PropertyGridVisible]
    [PropertyGridName("Smoothing")]
    public BezierPathSmoothing Smoothing
    {
        get
        {
            return smoothing;
        }
        set
        {
            if (smoothing == value)
            {
                return;
            }

            smoothing = value;
            if (Drawing != null)
            {
                RecalculateAndUpdate();
            }

            RaisePropertyChanged(nameof(Smoothing));
        }
    }

    public const double DefaultTension = 1;
    public const double MinimumTension = 0.5;
    public const double MaximumTension = 3;

    /// <summary>For a path being made: how it is smoothed, before it is worked out the first time</summary>
    public void InitializeSmoothing(BezierPathSmoothing smoothing, double tension)
    {
        this.smoothing = smoothing;
        this.tension = tension;
        curvesKnown = false;
    }

    /// <summary>
    /// How tight the automatic handles are: 1 as the method has it, more makes them shorter
    /// (bends tighter), less longer. Typed, or tied to a figure with a number (a slider),
    /// which it then follows (<see cref="ITiedValues"/>). Typed, from 0.5 to 3 (a slider in
    /// the grid, to try them out); a figure may give any number above 0.
    /// </summary>
    [PropertyGridVisible]
    [Domain(MinimumTension, MaximumTension)]
    [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
    public double Tension
    {
        get
        {
            switch (tensionSource)
            {
                case null:
                    return tension;
                case INumber number:
                    return number.Value;
                case ILengthProvider provider:
                    return provider.Length;
                default:
                    return double.NaN;
            }
        }
        set
        {
            if (tensionSource != null || tension == value)
            {
                return;
            }

            tension = value;
            if (Drawing != null)
            {
                RecalculateAndUpdate();
            }

            RaisePropertyChanged(nameof(Tension));
        }
    }

    /// <summary>The grid's button back from a tied tension (<see cref="Detach"/>)</summary>
    [PropertyGridVisible]
    [PropertyGridName("Type the tension")]
    [PropertyGridIcon(PropertyGridIcon.Pencil)]
    public void UntieTension()
    {
        if (Detach(nameof(Tension)))
        {
            Drawing.RaiseDisplayProperties(this);
        }
    }

    /// <summary>Every handle that is an offset from its anchor left to the path again (<see cref="Smoothing"/>), in one undo step</summary>
    [PropertyGridVisible]
    [PropertyGridName("Smooth all anchors")]
    [PropertyGridIcon(PropertyGridIcon.Arc)]
    public void SmoothAllAnchors()
    {
        if (Drawing == null || !CanSmoothAll)
        {
            return;
        }

        Drawing.ActionManager.RecordAction(LayoutChange(layout =>
        {
            foreach (var handle in layout.Handles.Keys.ToList())
            {
                if (layout.Handles[handle].Point == null)
                {
                    layout.Handles[handle] = HandleSpec.Automatic;
                }
            }
        }));
    }

    /// <summary>Whether a handle that bends a piece is the user's offset (not an image's handles, which are points)</summary>
    bool CanSmoothAll
    {
        get
        {
            return !IsImage && Handles.Any(h => h.Point == null && !h.Auto && BendsPiece(h, out _, out _));
        }
    }

    public bool CanEdit(string propertyName)
    {
        switch (propertyName)
        {
            case nameof(Tension):
                return tensionSource == null;
            case nameof(UntieTension):
                return this.IsTied(nameof(Tension));
            case nameof(SmoothAllAnchors):
                return CanSmoothAll;
            default:
                return true;
        }
    }

    /// <summary>A tension taken from a figure says which</summary>
    public string Caption(string propertyName, string defaultCaption)
    {
        if (propertyName == nameof(Tension) && tensionSource != null)
        {
            return "Tension = " + TiedValues.SourceName(tensionSource);
        }

        return defaultCaption;
    }

    #region Tied tension

    /// <summary>The tension, while a handle is worked out with it (none: no panel for it right after the tool)</summary>
    public IEnumerable<string> TiedValueNames
    {
        get
        {
            if ((smoothing != BezierPathSmoothing.None && Handles.Any(h => h.Auto)) || tensionSource != null)
            {
                yield return nameof(Tension);
            }
        }
    }

    /// <summary>The figure the tension comes from; null while it is typed</summary>
    public IFigure GetSource(string name)
    {
        return tensionSource;
    }

    /// <summary>
    /// A figure that says a number: a slider, a measurement, a label with a number - not
    /// what a point goes on (a segment): a click on one starts the next path
    /// </summary>
    public bool Accepts(string name, IFigure figure)
    {
        return !(figure is IPoint) && !(figure is ILinearFigure) && figure.GivesLength();
    }

    public bool TieTo(string name, IFigure source)
    {
        return TiedValues.Tie(
            new IFigure[] { this },
            tensionSource,
            source,
            owner => Drawing.ActionManager.RecordAction(LayoutChange(layout => layout.TensionSource = source)));
    }

    /// <summary>A typed tension again, at the value it has now</summary>
    public bool Detach(string name)
    {
        if (tensionSource == null || Drawing == null)
        {
            return false;
        }

        double value = Tension;
        Drawing.ActionManager.RecordAction(LayoutChange(layout =>
        {
            layout.TensionSource = null;
            layout.Tension = value > 0 && value.IsValidValue() ? value : DefaultTension;
        }));
        return true;
    }

    #endregion

    #region Anchors smooth or sharp

    /// <summary>The paths the point is an anchor of, whose handles can be changed (not images)</summary>
    static List<BezierPath> PathsWithAnchor(IFigure point)
    {
        if (point?.Drawing == null)
        {
            return new List<BezierPath>();
        }

        return point.Dependents
            .OfType<BezierPath>()
            .Distinct()
            .Where(path => path.Drawing != null && !path.IsImage && path.IsAnchor(point))
            .ToList();
    }

    /// <summary>The handles of the anchor that bend a piece</summary>
    IEnumerable<BezierPathHandle> BendingHandlesOf(IFigure anchor)
    {
        int index = Dependencies.IndexOf(anchor);
        if (index < 0 || index >= AnchorCount)
        {
            return Enumerable.Empty<BezierPathHandle>();
        }

        return new[] { inHandles[index], outHandles[index] }.Where(h => BendsPiece(h, out _, out _));
    }

    /// <summary>Whether a handle of the point, an anchor, isn't automatic</summary>
    public static bool CanSmoothAnchor(IFigure point)
    {
        return PathsWithAnchor(point).Any(path => path.BendingHandlesOf(point).Any(h => !h.Auto));
    }

    /// <summary>Whether a handle of the point, an anchor, is off it or automatic: it isn't a sharp corner yet</summary>
    public static bool CanSharpenAnchor(IFigure point)
    {
        return PathsWithAnchor(point).Any(path => path.BendingHandlesOf(point).Any(h => h.Auto || h.Point != null || h.Offset != default));
    }

    /// <summary>
    /// Both handles of the anchor left to its paths (<see cref="Smoothing"/>); one that was a
    /// point lets go of it. One undo step.
    /// </summary>
    public static void SmoothAnchor(IFigure point)
    {
        SetAnchorHandles(point, HandleSpec.Automatic);
    }

    /// <summary>Both handles of the anchor on it: the path turns there. One undo step.</summary>
    public static void SharpenAnchor(IFigure point)
    {
        SetAnchorHandles(point, default);
    }

    static void SetAnchorHandles(IFigure point, HandleSpec spec)
    {
        var paths = PathsWithAnchor(point);
        if (paths.Count == 0)
        {
            return;
        }

        var drawing = point.Drawing;
        using (Transaction.Create(drawing.ActionManager, false))
        {
            foreach (var path in paths)
            {
                drawing.ActionManager.RecordAction(path.LayoutChange(layout =>
                {
                    int index = layout.Anchors.IndexOf(point);
                    layout.Handles[layout.Ins[index]] = spec;
                    layout.Handles[layout.Outs[index]] = spec;
                }));
            }
        }
    }

    #endregion

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
            int count = System.Math.Min(AnchorCount, Dependencies.Count);
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
        int count = Dependencies.Count >= AnchorCount ? PieceCount : 0;
        int anchors = AnchorCount;
        if (count > 0)
        {
            SmoothHandles();
        }

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

    /// <summary>The automatic handles worked out from where the anchors and the other handles are now</summary>
    void SmoothHandles()
    {
        if (!Handles.Any(h => h.Auto))
        {
            return;
        }

        int count = AnchorCount;
        var points = new Point[count];
        for (int i = 0; i < count; i++)
        {
            points[i] = AnchorPoint(i);
            if (!points[i].Exists())
            {
                return;
            }
        }

        Point? Given(BezierPathHandle handle, int index)
        {
            return handle.Auto ? null : HandlePosition(handle, index).Minus(points[index]);
        }

        var (ins, outs) = BezierPathSmoother.Smooth(
            points,
            inHandles.Select((handle, i) => Given(handle, i)).ToList(),
            outHandles.Select((handle, i) => Given(handle, i)).ToList(),
            closed,
            smoothing,
            Tension);
        for (int i = 0; i < count; i++)
        {
            if (inHandles[i].Auto)
            {
                inHandles[i].Offset = ins[i];
            }

            if (outHandles[i].Auto)
            {
                outHandles[i].Offset = outs[i];
            }
        }
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

    /// <summary>
    /// The path exists while its anchors and the points of its handles do, and while the
    /// tension is a number above 0 if a handle is worked out with it; a hole that doesn't is
    /// left out
    /// </summary>
    public override void UpdateExistence()
    {
        int count = System.Math.Min(HoleStart, Dependencies.Count);
        bool exists = AnchorCount >= 2 && Dependencies.Count >= AnchorCount;
        for (int i = 0; i < count && exists; i++)
        {
            exists = Dependencies[i].Exists;
        }

        if (exists && smoothing != BezierPathSmoothing.None && Handles.Any(h => h.Auto))
        {
            double value = Tension;
            exists = (tensionSource == null || tensionSource.Exists) && value > 0 && value.IsValidValue();
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
                // a side that draws nothing (a transparent stroke: a letter's, a blob's) is
                // not what the click means: the filled inside is, within the cursor's reach of
                // its edge as any figure is - a thin shape was all edge, and every click on
                // it selected an invisible side
                if (!pieces[i].DrawsStroke)
                {
                    return filled && interior.Exists ? interior : null;
                }

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

    #region Points as handles

    /// <summary>
    /// Whether the point can be the handle's: not a handle itself, not the handle's own
    /// anchor, and not built on the path (it would be built on itself)
    /// </summary>
    public bool CanUseAsHandle(BezierPathHandle handle, IPoint point)
    {
        return !(point is BezierPathHandle)
            && point != AnchorOf(handle)
            && !point.DependsOn(this);
    }

    /// <summary>
    /// The handle is the point from now on (an Alt-drag of it dropped there): the path is
    /// built on the point, which comes before it in the list. Recorded, inside the caller's
    /// transaction.
    /// </summary>
    public void UsePointAsHandle(BezierPathHandle handle, IPoint point)
    {
        if (Drawing == null || !CanUseAsHandle(handle, point))
        {
            return;
        }

        Actions.MoveBefore(Drawing, point, this);
        Drawing.ActionManager.RecordAction(LayoutChange(layout =>
            layout.Handles[handle] = new HandleSpec(handle.Offset, point)));
    }

    #endregion

    #region Removing an anchor, a hole or the point of a handle

    /// <summary>
    /// A path keeps going without an anchor that is deleted, as long as two are left (not
    /// an image: its helper points would stay). A hole that is deleted leaves the inside
    /// whole again. A point that is a handle, deleted, leaves an ordinary handle where it was.
    /// The figure the tension comes from, deleted, leaves the tension typed.
    /// </summary>
    public bool CanRemoveDependency(IFigure dependency)
    {
        var indices = Enumerable.Range(0, Dependencies.Count).Where(i => Dependencies[i] == dependency).ToList();
        if (indices.Count == 0)
        {
            return false;
        }

        int anchors = AnchorCount;
        int holeStart = HoleStart;
        if (indices.All(i => i >= anchors && i < holeStart))
        {
            return true;
        }

        if (indices.Count != 1)
        {
            return false;
        }

        // two anchors left after the deletion - also of those that go with it (an anchor on
        // a segment through the one deleted): each asks before any of them has gone
        return indices[0] >= holeStart
            || indices[0] < anchors
                && Dependencies.Take(anchors).Count(anchor => !anchor.DependsOn(dependency)) >= 2
                && !IsImage;
    }

    public IAction GetRemoveDependencyAction(IFigure dependency)
    {
        int index = Dependencies.IndexOf(dependency);
        if (dependency == tensionSource && index == Dependencies.Count - 1)
        {
            // the tension stays what it was, typed
            double value = Tension;
            return LayoutChange(layout =>
            {
                layout.TensionSource = null;
                layout.Tension = value > 0 && value.IsValidValue() ? value : DefaultTension;
            });
        }

        if (index >= HoleStart)
        {
            return LayoutChange(layout => layout.Holes.Remove(dependency));
        }

        if (index >= AnchorCount)
        {
            return LayoutChange(layout => DetachPoint(layout, (IPoint)dependency));
        }

        return CreateRemoveAnchorAction(dependency, joinTarget: null);
    }

    /// <summary>Every handle that is the point is an ordinary handle where the point is</summary>
    void DetachPoint(Layout layout, IPoint point)
    {
        for (int i = 0; i < layout.Ins.Count; i++)
        {
            foreach (var handle in new[] { layout.Ins[i], layout.Outs[i] })
            {
                if (layout.Handles[handle].Point == point)
                {
                    var anchor = ((IPoint)layout.Anchors[i]).Coordinates;
                    layout.Handles[handle] = new HandleSpec(point.Coordinates.Minus(anchor), null);
                }
            }
        }
    }

    /// <summary>
    /// Whether the point can be dropped from the path into the target, an anchor next to it
    /// (an Alt-drag of one onto the other, <see cref="PointSnapping.Join"/>), with two
    /// anchors left
    /// </summary>
    public bool CanDropAnchorInto(IFigure point, IFigure target)
    {
        int count = AnchorCount;
        if (IsImage || HasImages || count <= 2)
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
        return LayoutChange(layout =>
        {
            int count = layout.Anchors.Count;
            int index = layout.Anchors.IndexOf(anchor);
            if (joinTarget != null)
            {
                int target = layout.Anchors.IndexOf(joinTarget);
                bool targetIsNext = target == (index + 1) % count;
                var heir = targetIsNext ? layout.Ins[target] : layout.Outs[target];
                var from = targetIsNext ? layout.Ins[index] : layout.Outs[index];
                layout.Handles[heir] = layout.Handles[from];
            }

            // the piece after the anchor goes (the one before, at the end of an open path),
            // and the one left joins its neighbors
            int pieceIndex = !closed && index == count - 1 ? index - 1 : index;
            MovePointsOff(layout, pieceIndex);
            layout.Pieces.RemoveAt(pieceIndex);
            layout.Ins.RemoveAt(index);
            layout.Outs.RemoveAt(index);
            layout.Anchors.RemoveAt(index);
        });
    }

    /// <summary>
    /// The points on the path are on the same pieces when the piece at the index goes: those
    /// after it a piece back, those on it onto the piece that takes its place, as far along it
    /// </summary>
    static void MovePointsOff(Layout layout, int pieceIndex)
    {
        foreach (var point in layout.Parameters.Keys.ToList())
        {
            var parameter = layout.Parameters[point];
            if (parameter >= pieceIndex + 1)
            {
                layout.Parameters[point] = parameter - 1;
            }
            else if (parameter >= pieceIndex)
            {
                layout.Parameters[point] = System.Math.Max(0, pieceIndex - 1) + (parameter - pieceIndex);
            }
        }
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
    /// a path (not an image), somewhere along a piece and not at its ends, and not one a
    /// locus is drawn from
    /// </summary>
    public static bool CanBecomeAnchor(IFigure point)
    {
        if (!(point is PointOnFigure onFigure)
            || !(onFigure.Dependencies.FirstOrDefault() is BezierPath path)
            || path.IsImage
            || path.HasImages
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
    /// the two pieces draw the same curve; a handle of the piece that is a point becomes an
    /// ordinary handle, since it moves). The points on the path stay where they are. One
    /// undo step.
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
        return LayoutChange(layout =>
        {
            int count = layout.Anchors.Count;
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

            // between two automatic handles the new anchor is smooth by itself too (the curve
            // changes a little); else the halves are the curve as it was
            var outHandle = layout.Outs[piece];
            var inHandle = layout.Ins[next];
            bool auto = layout.Handles[outHandle].Auto && layout.Handles[inHandle].Auto;
            if (!auto)
            {
                layout.Handles[outHandle] = new HandleSpec(p01.Minus(p0), null);
                layout.Handles[inHandle] = new HandleSpec(p23.Minus(p3), null);
            }

            var newIn = new BezierPathHandle(this, isIn: true) { Style = outHandle.Style };
            var newOut = new BezierPathHandle(this, isIn: false) { Style = outHandle.Style };
            var newPiece = new BezierPathPiece(this) { Style = layout.Pieces[piece].Style };
            layout.Handles[newIn] = auto ? HandleSpec.Automatic : new HandleSpec(p012.Minus(at), null);
            layout.Handles[newOut] = auto ? HandleSpec.Automatic : new HandleSpec(p123.Minus(at), null);
            layout.Ins.Insert(piece + 1, newIn);
            layout.Outs.Insert(piece + 1, newOut);
            layout.Pieces.Insert(piece + 1, newPiece);
            layout.Anchors.Insert(piece + 1, anchor);
            MovePointsAcross(layout, piece, t);
        });
    }

    static Point Lerp(Point from, Point to, double t)
    {
        return new Point(from.X + (to.X - from.X) * t, from.Y + (to.Y - from.Y) * t);
    }

    /// <summary>
    /// The points on the path stay where they are when the piece is split at t: those on
    /// later pieces a piece on, those on it onto the half they are on (the halves are the
    /// same cubic, run over [0, t] and [t, 1])
    /// </summary>
    static void MovePointsAcross(Layout layout, int piece, double t)
    {
        foreach (var point in layout.Parameters.Keys.ToList())
        {
            var parameter = layout.Parameters[point];
            if (parameter >= piece + 1)
            {
                layout.Parameters[point] = parameter + 1;
            }
            else if (parameter >= piece)
            {
                double along = parameter - piece;
                layout.Parameters[point] = along < t
                    ? piece + along / t
                    : piece + 1 + (along - t) / (1 - t);
            }
        }
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
            .Where(path => path != primary && !path.DependsOn(primary) && !primary.Holes.Contains(path))
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
                if (hole.Filled)
                {
                    Actions.SetProperty(drawing.ActionManager, new PropertyValue(nameof(Filled), hole), false);
                }
            }

            drawing.ActionManager.RecordAction(primary.LayoutChange(layout => layout.Holes.AddRange(holes)));
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
        var anchors = new List<IFigure>();
        for (int i = 0; i < count; i++)
        {
            anchors.Add(Transform(Dependencies[i], helper: false));
        }

        var ins = new List<HandleSpec>();
        var outs = new List<HandleSpec>();
        for (int i = 0; i < count; i++)
        {
            ins.Add(new HandleSpec(default, (IPoint)Transform(HandleSource(inHandles[i]), helper: true)));
            outs.Add(new HandleSpec(default, (IPoint)Transform(HandleSource(outHandles[i]), helper: true)));
        }

        var holes = Holes.Select(hole => Transform(hole, helper: false)).ToList();
        var path = Create(Drawing, anchors, ins, outs, holes, closed, filled);
        path.Visible = Visible;
        path.Style = Style;
        for (int i = 0; i < pieces.Count && i < path.pieces.Count; i++)
        {
            path.pieces[i].Style = pieces[i].Style;
        }

        result.Add(path);
        return result;
    }

    /// <summary>The handle as a point: its point, or the part itself</summary>
    static IFigure HandleSource(BezierPathHandle handle)
    {
        return (IFigure)handle.Point ?? handle;
    }

    /// <summary>
    /// Whether the transformations can take the path: what it is built on can all be
    /// transformed - but the figure the tension comes from, a number (the image's handles
    /// are points, which need none)
    /// </summary>
    public bool CanBeTransformed(Func<IFigure, bool> canBeTransformed)
    {
        return Dependencies.Where(dependency => dependency != tensionSource).All(canBeTransformed);
    }

    /// <summary>
    /// Whether something is built on the handles: the images of a transformation, whose
    /// helper points follow them. The path isn't split nor an anchor dropped then: the image
    /// would keep the old pieces.
    /// </summary>
    bool HasImages
    {
        get
        {
            return Handles.Any(handle => handle.Dependents.Any());
        }
    }

    #endregion

    #region File

    /// <summary>
    /// The anchors are the first dependencies, and their handles a path as Avalonia writes
    /// one, a piece from each anchor to the next, the closing one included whether or not the
    /// path is closed (it keeps its handles): "L" for a piece whose two handles are on their
    /// anchors, else "C", the first anchor's out handle and the second one's in handle - each
    /// an offset from its anchor ("1,0.5"), a point of the drawing, by its place among the
    /// dependencies ("#4"), or "a", automatic. The dependencies after those are the holes,
    /// and last the figure the tension comes from, when it does: <c>Tension="#7"</c>, where a
    /// typed one is a number.
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

        if (smoothing != BezierPathSmoothing.Hobby)
        {
            writer.WriteAttributeString("Smoothing", smoothing.ToString());
        }

        if (tensionSource != null)
        {
            writer.WriteAttributeString("Tension", "#" + (Dependencies.Count - 1).ToString(CultureInfo.InvariantCulture));
        }
        else if (tension != DefaultTension)
        {
            writer.WriteAttributeDouble("Tension", tension);
        }

        writer.WriteAttributeString("Path", HandlesText());
        WritePartStyles(writer);
    }

    string HandlesText()
    {
        // where each point of a handle is among the dependencies
        var places = new Dictionary<BezierPathHandle, int>();
        int place = AnchorCount;
        foreach (var handle in HandlesInOrder.Where(h => h.Point != null))
        {
            places[handle] = place++;
        }

        string Write(BezierPathHandle handle)
        {
            if (handle.Auto)
            {
                return AutomaticToken;
            }

            return handle.Point != null
                ? "#" + places[handle].ToString(CultureInfo.InvariantCulture)
                : Number(handle.Offset.X) + "," + Number(handle.Offset.Y);
        }

        var text = new StringBuilder();
        int count = inHandles.Count;
        for (int i = 0; i < count; i++)
        {
            if (i > 0)
            {
                text.Append(' ');
            }

            var outHandle = outHandles[i];
            var inHandle = inHandles[(i + 1) % count];
            if (outHandle.Point == null
                && inHandle.Point == null
                && !outHandle.Auto
                && !inHandle.Auto
                && outHandle.Offset == default
                && inHandle.Offset == default)
            {
                text.Append('L');
            }
            else
            {
                text.Append("C ").Append(Write(outHandle)).Append(' ').Append(Write(inHandle));
            }
        }

        return text.ToString();
    }

    const string AutomaticToken = "a";

    static string Number(double value)
    {
        // (plus zero: a negative zero would be written "-0")
        return (value + 0.0).ToStringInvariant();
    }

    /// <summary>
    /// The path <see cref="HandlesText"/> wrote, over the dependencies the file gave: what
    /// can't be read leaves a handle on its anchor; with no path at all, the points the
    /// dependencies start with are the anchors
    /// </summary>
    public override void ReadXml(XElement element)
    {
        closed = element.ReadBool("Closed", false);
        filled = element.ReadBool("Filled", false);
        var listed = Dependencies.ToList();
        var tokens = (element.ReadString("Path") ?? "").Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
        int count = tokens.Count(token => token == "L" || token == "C");
        if (count < 2 || count > listed.Count || !listed.Take(count).All(d => d is IPoint))
        {
            count = listed.TakeWhile(d => d is IPoint).Count();
            tokens = Array.Empty<string>();
        }

        var ins = Enumerable.Repeat(default(HandleSpec), count).ToList();
        var outs = Enumerable.Repeat(default(HandleSpec), count).ToList();
        var referenced = new HashSet<int>();
        HandleSpec Read(string text)
        {
            if (text == AutomaticToken)
            {
                return HandleSpec.Automatic;
            }

            if (text.StartsWith("#", StringComparison.Ordinal)
                && int.TryParse(text.Substring(1), NumberStyles.None, CultureInfo.InvariantCulture, out int index)
                && index >= count
                && index < listed.Count
                && listed[index] is IPoint point)
            {
                referenced.Add(index);
                return new HandleSpec(default, point);
            }

            return new HandleSpec(TryParsePoint(text, out var offset) ? offset : default, null);
        }

        int piece = 0;
        for (int i = 0; i < tokens.Length && piece < count; piece++)
        {
            if (tokens[i] == "C" && i + 2 < tokens.Length)
            {
                outs[piece] = Read(tokens[i + 1]);
                ins[(piece + 1) % count] = Read(tokens[i + 2]);
                i += 3;
            }
            else
            {
                i++;
            }
        }

        smoothing = Enum.TryParse(element.ReadString("Smoothing"), out BezierPathSmoothing read) ? read : BezierPathSmoothing.Hobby;
        tension = DefaultTension;
        tensionSource = null;
        var tensionText = element.ReadString("Tension");
        if (tensionText != null
            && tensionText.StartsWith("#", StringComparison.Ordinal)
            && int.TryParse(tensionText.Substring(1), NumberStyles.None, CultureInfo.InvariantCulture, out int tensionIndex)
            && tensionIndex >= count
            && tensionIndex < listed.Count)
        {
            tensionSource = listed[tensionIndex];
            referenced.Add(tensionIndex);
        }
        else if (double.TryParse(tensionText, NumberStyles.Float, CultureInfo.InvariantCulture, out double typed) && typed.IsValidValue())
        {
            tension = typed;
        }

        var holes = Enumerable.Range(count, listed.Count - count)
            .Where(index => !referenced.Contains(index) && listed[index] is BezierPath)
            .Select(index => listed[index])
            .ToList();
        Build(listed.Take(count).ToList(), ins, outs, holes);

        // visibility and the style of the inside, onto the parts made
        base.ReadXml(element);
        ReadPartStyles(element);
        curvesKnown = false;
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
    /// A handle of an anchor: an offset from it, which the Drag tool changes, or a point of
    /// the drawing (<see cref="Point"/>; the part is then hidden, the point shows itself).
    /// Not selected by itself (a click on it leaves the selection as it is, so that it stays
    /// shown); styled with the others on the path's page.
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

        /// <summary>Where the handle is from its anchor, unless it is a point</summary>
        public Point Offset { get; set; }

        /// <summary>The point of the drawing the handle is, if it is one</summary>
        public IPoint Point { get; set; }

        /// <summary>
        /// Whether the path works the handle out from its anchors (<see cref="BezierPath.Smoothing"/>);
        /// <see cref="Offset"/> is then what it came to last. A drag makes it the user's.
        /// </summary>
        public bool Auto { get; set; }

        public HandleSpec Spec
        {
            get
            {
                return new HandleSpec(Offset, Point, Auto);
            }
        }

        public void Set(HandleSpec spec)
        {
            Offset = spec.Offset;
            Point = spec.Point;
            Auto = spec.Auto && spec.Point == null;
        }

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
        protected override ZOrder DefaultLayer()
        {
            return ZOrder.Handles;
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
            return !Owner.Locked && Owner.Drawing != null && Point == null;
        }

        public override void MoveToCore(Avalonia.Point newLocation)
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

        protected override ZOrder DefaultLayer()
        {
            return ZOrder.Figures;
        }

        // the path finds its parts (BezierPath.HitTest)
        public override IFigure HitTest(Point point)
        {
            return Owner.HitTest(point) == this ? this : null;
        }

        /// <summary>Whether the side shows at all: a stroke of some width in a color that is not fully transparent</summary>
        public bool DrawsStroke
        {
            get
            {
                return Shape.StrokeThickness > 0
                    && Shape.Stroke != null
                    && !(Shape.Stroke is ISolidColorBrush brush && brush.Color.A == 0);
            }
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

        protected override ZOrder DefaultLayer()
        {
            return ZOrder.Polygons;
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
