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
    bool isQueued;

    public DrawingThumbnail(GalleryItem item)
    {
        this.item = item;
        Stretch = Stretch.Uniform;
        IsHitTestVisible = false;
        Child = surface;
        surface.SizeChanged += (s, e) => Enqueue();
    }

    public Drawing Drawing { get; private set; }

    /// <summary>The drawing is on the surface (the tile takes its paper from it)</summary>
    public event Action<Drawing> Loaded = delegate { };

    // Drawings are loaded one at a time when the UI thread has nothing better to do: the
    // gallery shows up at once and fills in, first tile first.
    static readonly Queue<DrawingThumbnail> queue = new Queue<DrawingThumbnail>();
    static bool isPumping;

    void Enqueue()
    {
        if (isQueued || surface.Bounds.Width <= 0)
        {
            return;
        }

        isQueued = true;
        queue.Enqueue(this);
        Pump();
    }

    static void Pump()
    {
        if (isPumping || queue.Count == 0)
        {
            return;
        }

        isPumping = true;
        Dispatcher.UIThread.Post(
            () =>
            {
                isPumping = false;
                if (queue.Count > 0)
                {
                    queue.Dequeue().Load();
                }

                Pump();
            },
            DispatcherPriority.Background);
    }

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

    void Load()
    {
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
            surface.Background = null;
            drawing.CoordinateSystem.ZoomExtend(item.Plane);
            drawing.CoordinateSystem.Zoom(zoomAfterFit, new Point(SurfaceWidth / 2, SurfaceHeight / 2));
            Drawing = drawing;
            Loaded(drawing);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Gallery: " + item.FileName + ": " + ex);
        }
    }
}
