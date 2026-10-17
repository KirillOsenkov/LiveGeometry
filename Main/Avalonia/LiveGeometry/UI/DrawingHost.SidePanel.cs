using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;

namespace LiveGeometry;

/// <summary>
/// Where the side panel goes, so that it doesn't cover what it is about. Beside the drawing
/// in a wide window, at the right or at the left as the user chose (the button by its
/// cross); across a narrow one (a phone), at most half its height, at the top or the bottom.
/// What was picked in the plane and is under the panel is brought out by moving the view,
/// which goes back when the panel goes. What the view can't move - text pinned to the
/// screen, a caption - or what has no room to go to, gets the panel on the other side.
/// </summary>
public partial class DrawingHost
{
    enum PanelSide
    {
        Right,
        Left,
        Top,
        Bottom
    }

    PanelSide panelSide = PanelSide.Right;

    // set while the panel is shown for a selection just made (not for undo, nor a tool's panel)
    bool revealSelection;

    // the button by the panel's cross that moves it to the other side
    Border sideButton;

    bool panelAtLeft;

    /// <summary>
    /// Whether the panel is at the left of a wide window rather than the right: the user's
    /// choice (the button by its cross), which the app keeps between runs. The panel still
    /// goes to the other side for a while when what it is about is under it and the view
    /// can't move it out.
    /// </summary>
    public bool PanelAtLeft
    {
        get => panelAtLeft;
        set
        {
            panelAtLeft = value;
            PlaceSidePanel();
        }
    }

    /// <summary>The user moved the panel to the other side (<see cref="PanelAtLeft"/>), to be kept</summary>
    public event Action PanelSideChosen;

    // set while the panel goes to the side the user just chose: it stays there
    bool sideChosen;

    void MovePanelToOtherSide()
    {
        panelAtLeft = panelSide != PanelSide.Left;
        sideChosen = true;
        revealSelection = true;
        try
        {
            PlaceSidePanel();
        }
        finally
        {
            sideChosen = false;
            revealSelection = false;
        }

        PanelSideChosen?.Invoke();
    }

    /// <summary>An arrow to the edge the panel would go to</summary>
    static string SideButtonPath(bool toLeft)
    {
        return toLeft
            ? "M0.75,0 L0.75,9 M10,4.5 L3.5,4.5 M6,2 L3.5,4.5 L6,7"
            : "M9.25,0 L9.25,9 M0,4.5 L6.5,4.5 M4,2 L6.5,4.5 L4,7";
    }

    /// <summary>The room the drawing keeps beside the panel, in pixels; with less the panel goes across</summary>
    const double MinimumRoomBeside = 240;

    /// <summary>The share of the height a panel across the drawing takes at most</summary>
    const double StackedShare = 0.5;

    const double PanelMargin = 8;

    /// <summary>Between the panel and what was brought out from under it: room for a point's name</summary>
    const double RevealGap = 32;

    // the placement waiting for the page to settle, if any
    int pendingPlacement;

    /// <summary>
    /// The panel was given a page to show, or put away. A page fills in a moment after it is
    /// built (the value under an expression box, the swatches of a style), and only then has
    /// its size: the panel is placed then, and stays unseen until it is, so that it never
    /// shows on the wrong side first.
    /// </summary>
    void PlaceSidePanel()
    {
        int placement = ++pendingPlacement;
        if (!sidePanel.IsVisible)
        {
            sidePanel.Opacity = 1;
            PlaceSidePanelNow(reveal: false);
            return;
        }

        sidePanel.Opacity = 0;
        bool reveal = revealSelection;
        bool chosen = sideChosen;
        Dispatcher.UIThread.Post(
            () =>
            {
                if (placement == pendingPlacement)
                {
                    sidePanel.Opacity = 1;
                    PlaceSidePanelNow(reveal, chosen);
                }
            },
            DispatcherPriority.Background);
    }

    /// <param name="reveal">The page is about a selection just made: the view may move to show it</param>
    /// <param name="chosen">The user just chose the side: the panel stays on it</param>
    void PlaceSidePanelNow(bool reveal, bool chosen = false)
    {
        var drawing = CurrentDrawing;
        var canvasSize = DrawingControl.Bounds.Size;
        if (!sidePanel.IsVisible || drawing == null || canvasSize.Width <= 0 || canvasSize.Height <= 0)
        {
            // (the panel is away while a text is edited where it is, and comes back after)
            if (InPlaceTextEditor.Current == null)
            {
                RestoreRevealedView();
            }

            return;
        }

        // Across a window with no room beside the panel (a phone): first its width beside
        // the drawing. Laid out for real: measured alone, the rows' shared caption column
        // has not settled, and the panel came out narrower than it is.
        ApplySide(PanelSide.Right, stacked: false);
        sidePanel.MaxHeight = double.PositiveInfinity;
        propertyGridScrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
        sidePanel.UpdateLayout();
        bool stacked = canvasSize.Width < sidePanel.Bounds.Width + 2 * PanelMargin + MinimumRoomBeside;
        if (stacked)
        {
            // the rows wrap to the width of the screen, and scroll under half its height
            ApplySide(PanelSide.Top, stacked: true);
            sidePanel.MaxHeight = System.Math.Max(0, canvasSize.Height * StackedShare - 2 * PanelMargin);
            propertyGridScrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
            sidePanel.UpdateLayout();
        }

        // (across a phone there is no other side to choose)
        sideButton.IsVisible = !stacked;
        var defaultSide = stacked ? PanelSide.Top : panelAtLeft ? PanelSide.Left : PanelSide.Right;
        var otherSide = stacked ? PanelSide.Bottom : panelAtLeft ? PanelSide.Right : PanelSide.Left;
        var size = sidePanel.Bounds.Size;
        var shown = PropertyGrid.Selection;
        if (!(shown is IFigure || shown is FigureSelection))
        {
            ApplySide(defaultSide, stacked);
            return;
        }

        // what was picked: in the plane, or pinned to the screen (measured where a move of the
        // view under way would leave it)
        viewGlide?.Finish();
        var canvas = drawing.Canvas;
        Rect? plane = null;
        Rect? pinned = null;
        foreach (var figure in drawing.GetSelectedFigures())
        {
            var bounds = SelectionRegion.FindShownBounds(figure, canvas);
            if (bounds == null)
            {
                continue;
            }

            if (IsPinned(figure))
            {
                pinned = pinned?.Union(bounds.Value) ?? bounds.Value;
            }
            else
            {
                plane = plane?.Union(bounds.Value) ?? bounds.Value;
            }
        }

        double Covered(PanelSide candidate, Rect? bounds) => bounds == null ? 0 : Area(PanelRect(candidate, size, canvasSize).Intersect(bounds.Value));

        // the text pinned to the screen, which the view can't move out from under (a caption
        // at the bottom of a phone)
        var obstacles = drawing.Figures
            .Where(figure => figure.Visible && !figure.Selected && IsPinned(figure))
            .Select(figure => SelectionRegion.FindShownBounds(figure, canvas))
            .Where(bounds => bounds != null)
            .Select(bounds => bounds.Value)
            .ToList();

        // The other side, for this page only, for what the view can't bring out: text pinned
        // to the screen, or a figure with no room to go to. Not when the user just chose
        // the side: then only the view moves.
        var side = defaultSide;
        if (!chosen
            && ((pinned != null && Covered(defaultSide, pinned) > 0 && Covered(otherSide, pinned) < Covered(defaultSide, pinned))
                || (plane != null
                    && FindRevealMove(plane.Value, PanelRect(defaultSide, size, canvasSize), canvasSize, obstacles) == null
                    && Covered(otherSide, plane) < Covered(defaultSide, plane))))
        {
            side = otherSide;
        }

        // the side every time the page is about figures, but the view only moves for what
        // was just picked: after undo it stays where it is
        ApplySide(side, stacked);
        if (plane != null && reveal)
        {
            var move = FindRevealMove(plane.Value, PanelRect(side, size, canvasSize), canvasSize, obstacles);
            if (move != null && move.Value != default)
            {
                RevealInPlane(move.Value);
            }
        }
    }

    /// <summary>Text pinned to the screen, which stays put while the view moves</summary>
    static bool IsPinned(IFigure figure)
    {
        return figure is Label { Pin: not LabelPin.None } or ShowHideControl { Pin: not LabelPin.None };
    }

    static Rect PanelRect(PanelSide side, Size size, Size canvas)
    {
        return side switch
        {
            PanelSide.Left => new Rect(PanelMargin, PanelMargin, size.Width, size.Height),
            PanelSide.Bottom => new Rect(PanelMargin, canvas.Height - PanelMargin - size.Height, size.Width, size.Height),
            PanelSide.Top => new Rect(PanelMargin, PanelMargin, size.Width, size.Height),
            _ => new Rect(canvas.Width - PanelMargin - size.Width, PanelMargin, size.Width, size.Height)
        };
    }

    void ApplySide(PanelSide side, bool stacked)
    {
        panelSide = side;
        sidePanel.HorizontalAlignment = stacked
            ? HorizontalAlignment.Stretch
            : side == PanelSide.Left ? HorizontalAlignment.Left : HorizontalAlignment.Right;
        sidePanel.VerticalAlignment = side == PanelSide.Bottom ? VerticalAlignment.Bottom : VerticalAlignment.Top;

        // the button says where it would take the panel
        bool toLeft = side != PanelSide.Left;
        ((Avalonia.Controls.Shapes.Path)sideButton.Child).Data = Geometry.Parse(SideButtonPath(toLeft));
        ToolTip.SetTip(sideButton, toLeft ? "Move the panel to the left" : "Move the panel to the right");
    }

    static double Area(Rect rect) => rect.Width > 0 && rect.Height > 0 ? rect.Width * rect.Height : 0;

    #region Moving the view

    // how far the view was moved to bring the selection out, and where it was left: the
    // view goes back by as much when the panel goes, unless it was moved since
    Avalonia.Vector revealedBy;
    Point viewAfterReveal;
    double zoomAfterReveal;

    Glide viewGlide;

    /// <summary>
    /// The least move of the view that brings what was picked out from under the panel, whole,
    /// on the canvas and clear of the text pinned to the screen: none (zero) when it is not
    /// under the panel, null when no move does it
    /// </summary>
    static Avalonia.Vector? FindRevealMove(Rect bounds, Rect panel, Size canvas, List<Rect> obstacles)
    {
        bounds = bounds.Inflate(RevealGap);
        if (!bounds.Intersects(panel))
        {
            return default(Avalonia.Vector);
        }

        var room = new Rect(canvas);
        var candidates = new[]
        {
            new Avalonia.Vector(panel.Left - bounds.Right, 0),
            new Avalonia.Vector(panel.Right - bounds.Left, 0),
            new Avalonia.Vector(0, panel.Bottom - bounds.Top),
            new Avalonia.Vector(0, panel.Top - bounds.Bottom)
        };
        return candidates
            .Where(candidate =>
            {
                var moved = bounds.Translate(candidate);
                return room.Contains(moved) && !moved.Intersects(panel) && !obstacles.Any(moved.Intersects);
            })
            .OrderBy(candidate => candidate.Length)
            .Cast<Avalonia.Vector?>()
            .FirstOrDefault();
    }

    /// <summary>
    /// Moves the view to bring what was picked out from under the panel (gliding there). Not
    /// an undo step: the app moved the view to show what was picked, as a fit to a resized
    /// window does, and undo after a click would first undo a move the user never made. (A
    /// view's move in the history is an offset, so one left out of it upsets nothing.)
    /// </summary>
    void RevealInPlane(Avalonia.Vector move)
    {
        var coordinateSystem = CurrentDrawing.CoordinateSystem;
        viewGlide?.Finish();
        if (!IsViewAsRevealed())
        {
            revealedBy = default;
        }

        revealedBy += move;
        MoveViewSmoothly(move, () =>
        {
            viewAfterReveal = coordinateSystem.Origin;
            zoomAfterReveal = coordinateSystem.UnitLength;
        });
    }

    bool IsViewAsRevealed()
    {
        var coordinateSystem = CurrentDrawing?.CoordinateSystem;
        return coordinateSystem != null
            && coordinateSystem.Origin == viewAfterReveal
            && coordinateSystem.UnitLength == zoomAfterReveal;
    }

    /// <summary>The panel went: the view goes back to where it was, if nothing moved it since</summary>
    void RestoreRevealedView()
    {
        viewGlide?.Finish();
        if (revealedBy == default(Avalonia.Vector) || !IsViewAsRevealed())
        {
            revealedBy = default;
            return;
        }

        var back = -revealedBy;
        revealedBy = default;
        MoveViewSmoothly(back, null);
    }

    void MoveViewSmoothly(Avalonia.Vector move, Action done)
    {
        viewGlide?.Finish();
        var drawing = CurrentDrawing;
        var coordinateSystem = drawing.CoordinateSystem;
        var start = coordinateSystem.Origin;
        viewGlide = new Glide(progress =>
        {
            if (drawing != CurrentDrawing)
            {
                return;
            }

            coordinateSystem.MoveTo(coordinateSystem.ToLogical(start + move * progress));
            if (progress >= 1)
            {
                done?.Invoke();
            }
        });
    }

    /// <summary>A short eased movement: the step is told how far along it is, from 0 to 1</summary>
    class Glide
    {
        static readonly TimeSpan duration = TimeSpan.FromMilliseconds(220);

        readonly Action<double> step;
        readonly Stopwatch watch = Stopwatch.StartNew();
        DispatcherTimer timer;

        public Glide(Action<double> step)
        {
            this.step = step;
            timer = new DispatcherTimer(TimeSpan.FromMilliseconds(15), DispatcherPriority.Render, (s, e) => Tick());
            timer.Start();
            Tick();
        }

        void Tick()
        {
            double t = System.Math.Min(1, watch.Elapsed.TotalMilliseconds / duration.TotalMilliseconds);
            if (t >= 1)
            {
                Finish();
                return;
            }

            // ease out: quick at first, settling at the end
            step(1 - System.Math.Pow(1 - t, 3));
        }

        /// <summary>To the end at once</summary>
        public void Finish()
        {
            if (timer == null)
            {
                return;
            }

            timer.Stop();
            timer = null;
            step(1);
        }
    }

    #endregion
}
