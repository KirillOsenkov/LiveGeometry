using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using DynamicGeometry;
using Drawing = DynamicGeometry.Drawing;

namespace LiveGeometry;

/// <summary>
/// A small picture of a drawing that is not a picture: the real drawing on a canvas of its own,
/// laid out at <see cref="SurfaceWidth"/> x <see cref="SurfaceHeight"/> and scaled down to
/// whatever size the thumbnail gets, so that points, strokes and text shrink together. Nothing
/// to generate, nothing to go stale, and it looks exactly like the drawing will when opened.
/// No tool is attached to the drawing, so it doesn't react to the mouse.
/// </summary>
public class DrawingThumbnail : Viewbox
{
    public const double SurfaceWidth = 560;
    public const double SurfaceHeight = 380;

    // what zoom to fit leaves around the content is meant for a full window
    const double zoomAfterFit = 1.08;

    readonly Canvas surface = new Canvas()
    {
        Width = SurfaceWidth,
        Height = SurfaceHeight,
        ClipToBounds = true,
        IsHitTestVisible = false
    };

    readonly GalleryItem item;

    public DrawingThumbnail(GalleryItem item)
    {
        this.item = item;
        Stretch = Stretch.Uniform;
        IsHitTestVisible = false;
        Child = surface;

        // the figures follow the theme like the editor's do, while the tile is on screen (the
        // tiles live as long as the gallery); hidden behind the editor, a tile catches up when
        // the gallery page shows again (RefreshThemeIfStale, called by the gallery)
        AppTheme.CurrentChanged += () => RefreshTheme(colorsChanged: false);
        AppTheme.ColorsChanged += () => RefreshTheme(colorsChanged: true);
    }

    void RefreshTheme(bool colorsChanged)
    {
        if (Drawing != null && IsEffectivelyVisible)
        {
            Drawing.RefreshTheme(colorsChanged);
        }
    }

    /// <summary>The theme may have changed while the tile was hidden</summary>
    public void RefreshThemeIfStale()
    {
        Drawing?.RefreshThemeIfStale();
    }

    public Drawing Drawing { get; private set; }

    /// <summary>The drawing is on the surface (the tile takes its paper from it)</summary>
    public event Action<Drawing> DrawingLoaded = delegate { };

    // Load has been called: the drawing is on the surface, or it failed to load and the
    // console says why (it is not tried again)
    bool isLoadCalled;

    /// <summary>
    /// Not loaded yet, and the surface has its size: the drawing would be fitted to it. The
    /// gallery decides when (<see cref="GalleryView"/>); until then the tile is its plate and
    /// its caption.
    /// </summary>
    public bool CanLoad => !isLoadCalled && surface.Bounds.Width > 0;

    /// <summary>See <see cref="GalleryItem.UsesEmoji"/></summary>
    public bool UsesEmoji => item.UsesEmoji;

    #region Animation

    // While the mouse is over the tile, the points that can be dragged drift around in small
    // circles and the construction follows: a still picture can't say "this is alive".

    const double driftPixels = 16;
    const double driftSeconds = 3.2;
    static readonly TimeSpan frame = TimeSpan.FromMilliseconds(33);

    class Drifter
    {
        public IMovable Point;
        public Point Home;
        public double Phase;
        public double Direction;
    }

    readonly List<Drifter> drifters = new List<Drifter>();
    readonly Stopwatch clock = new Stopwatch();
    DispatcherTimer timer;

    public bool IsAnimated
    {
        get => timer != null;
        set
        {
            if (value == IsAnimated || Drawing == null)
            {
                return;
            }

            if (value)
            {
                StartDrifting();
            }
            else
            {
                StopDrifting();
            }
        }
    }

    void StartDrifting()
    {
        drifters.Clear();
        foreach (var figure in Drawing.Figures)
        {
            if (figure.Visible && figure is IPoint && figure is IMovable movable && movable.AllowMove())
            {
                int index = drifters.Count;
                drifters.Add(new Drifter()
                {
                    Point = movable,
                    Home = movable.Coordinates,
                    Phase = index * 2.4,
                    Direction = index % 2 == 0 ? 1 : -1
                });
            }
        }

        if (drifters.Count == 0)
        {
            return;
        }

        clock.Restart();
        timer = new DispatcherTimer() { Interval = frame };
        timer.Tick += (s, e) => Drift(clock.Elapsed.TotalSeconds);
        timer.Start();
    }

    void StopDrifting()
    {
        timer.Stop();
        timer = null;
        foreach (var drifter in drifters)
        {
            drifter.Point.MoveTo(drifter.Home);
        }

        Drawing.Recalculate();
    }

    void Drift(double seconds)
    {
        try
        {
            // starts and ends at home: no jump when the mouse comes and goes
            double radius = Drawing.CoordinateSystem.ToLogical(driftPixels);
            double turn = 2 * System.Math.PI * seconds / driftSeconds;
            foreach (var drifter in drifters)
            {
                double angle = drifter.Direction * turn + drifter.Phase;
                drifter.Point.MoveTo(new Point(
                    drifter.Home.X + radius * (System.Math.Cos(angle) - System.Math.Cos(drifter.Phase)),
                    drifter.Home.Y + radius * (System.Math.Sin(angle) - System.Math.Sin(drifter.Phase))));
            }

            Drawing.Recalculate();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Gallery: " + item.FileName + ": " + ex.Message);
            StopDrifting();
        }
    }

    #endregion

    public void Load()
    {
        isLoadCalled = true;
        try
        {
            var drawing = new Drawing(surface);
            drawing.Status += text =>
            {
                if (!string.IsNullOrEmpty(text))
                {
                    Console.WriteLine("Gallery: " + item.FileName + ": " + text);
                }
            };

            PointBase.SuppressAutoLabelPoints = true;
            try
            {
                drawing.AddFromXml(XElement.Parse(item.LoadText()));
            }
            finally
            {
                PointBase.SuppressAutoLabelPoints = false;
            }

            GalleryDrawing.HideText(drawing);

            // the tile shows through: the drawing's paper, if it has one, is the tile's plate
            drawing.PaintsPaper = false;
            surface.Background = null;
            var scene = drawing.ChooseScene(SurfaceWidth, SurfaceHeight);
            if (scene != null)
            {
                drawing.ShowScene(scene.Value);
            }
            else
            {
                drawing.CoordinateSystem.ZoomExtend(item.Plane);
                drawing.CoordinateSystem.Zoom(zoomAfterFit, new Point(SurfaceWidth / 2, SurfaceHeight / 2));
            }

            Drawing = drawing;
            DrawingLoaded(drawing);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Gallery: " + item.FileName + ": " + ex);
        }
    }
}
