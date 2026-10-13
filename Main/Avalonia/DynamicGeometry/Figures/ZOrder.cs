using System;
using System.Collections.Generic;
using System.Linq;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    /// <summary>
    /// The layers figures are drawn in, bottom up. A figure's <see cref="IFigure.Layer"/> is
    /// its kind's. Two bands of layers are ordered by <see cref="IFigure.Z"/> (0 unless Bring
    /// to front or Send to back changed it), the Z first and the layer second
    /// (<see cref="ZOrders.Encode"/>): the figures', from <see cref="Polygons"/> to
    /// <see cref="Vectors"/>, so a polygon brought to front covers a circle; and the points',
    /// from <see cref="FixedPoints"/> to <see cref="FreePoints"/>, all of it over the first,
    /// so nothing in the figures' band ever covers a point - free points over those dragged
    /// along a figure over those that don't move by themselves, and any of them brought to
    /// front over all three. Within one value the later figure is on top.
    /// </summary>
    public enum ZOrder : int
    {
        Grid = 10,
        Axes = 20,
        Polygons = 30,
        Labels = 40,
        Figures = 50,
        Vectors = 55,
        SelectionHalos = 58,
        Handles = 59,

        /// <summary>Points that a drag doesn't move by themselves: midpoints, intersections, images, points by coordinates</summary>
        FixedPoints = 60,

        /// <summary>Points dragged along a figure or within one: on a figure, inside a shape, translated with something free; sliders</summary>
        ConstrainedPoints = 62,

        /// <summary>Points dragged anywhere</summary>
        FreePoints = 64,

        PointLabels = 70,
        Controls = 80,

        StatusBar = 100,
        PropertyGrid = 110,
    }

    /// <summary>
    /// What a figure's layer and Z make of its ZIndex, and the two verbs that change the Z
    /// (<see cref="ZOrder"/>).
    /// </summary>
    public static class ZOrders
    {
        /// <summary>How far a Z may go either way; the verbs only ever step by one past the extremes</summary>
        public const int Range = 999;

        /// <summary>One step of Z is this many ZIndex values: room for every layer of the band</summary>
        public const int Stride = 1000;

        /// <summary>All the Z values of a band, with their layers</summary>
        const int Span = Stride * (2 * Range + 1);

        /// <summary>Where the figures' band starts: over the grid and the axes</summary>
        const int FigureBand = Stride;

        /// <summary>Where the points' band starts: over the figures' band and the selection halos and handles between the two</summary>
        const int PointBand = FigureBand + Span + Stride;

        /// <summary>
        /// The band whose Z orders figures of the layer, named by its lowest layer: the
        /// figures' (<see cref="ZOrder.Polygons"/> to <see cref="ZOrder.Vectors"/>), the
        /// points' (<see cref="ZOrder.FixedPoints"/> to <see cref="ZOrder.FreePoints"/>), or none
        /// </summary>
        static ZOrder? Band(ZOrder layer)
        {
            if (layer >= ZOrder.Polygons && layer <= ZOrder.Vectors)
            {
                return ZOrder.Polygons;
            }

            if (layer >= ZOrder.FixedPoints && layer <= ZOrder.FreePoints)
            {
                return ZOrder.FixedPoints;
            }

            return null;
        }

        /// <summary>Whether a figure of the layer can be brought to front and sent to back: it is in one of the bands</summary>
        public static bool IsMovable(ZOrder layer)
        {
            return Band(layer) != null;
        }

        /// <summary>Whether the Z of figures of the two layers orders them among each other: a Z is a place in its own band only</summary>
        public static bool AreInSameBand(ZOrder layer, ZOrder other)
        {
            return IsMovable(layer) && Band(layer) == Band(other);
        }

        /// <summary>
        /// The ZIndex of the layer at the Z, bottom up: the layers below the figures' band as
        /// they are, the figures' band ordered by Z and then by layer, the selection halos and
        /// the handles, the points' band ordered the same way, the layers above it over all of it.
        /// </summary>
        public static int Encode(ZOrder layer, int z)
        {
            if (layer < ZOrder.Polygons)
            {
                return (int)layer;
            }

            if (layer <= ZOrder.Vectors)
            {
                return FigureBand + Place(z) + (int)layer;
            }

            if (layer < ZOrder.FixedPoints)
            {
                return FigureBand + Span + (int)layer;
            }

            if (layer <= ZOrder.FreePoints)
            {
                return PointBand + Place(z) + (int)layer;
            }

            return PointBand + Span + (int)layer;
        }

        /// <summary>
        /// The Z's part of a ZIndex in its band: a stride per step, so that what is drawn just
        /// under a figure (a halo, at its ZIndex - 1) is still over the Z below
        /// </summary>
        static int Place(int z)
        {
            return (System.Math.Clamp(z, -Range, Range) + Range) * Stride;
        }

        /// <summary>The ZIndex of a figure of the layer whose Z was never changed</summary>
        public static int Default(ZOrder layer)
        {
            return Encode(layer, 0);
        }

        /// <summary>
        /// A part is drawn where its figure is: it takes the figure's Z when the two are in one
        /// band, and keeps its own otherwise (a regular polygon's vertices, a slider's track)
        /// </summary>
        public static void PassOn(IFigure figure, IFigure part)
        {
            if (AreInSameBand(figure.Layer, part.Layer))
            {
                part.Z = figure.Z;
            }
        }

        /// <summary>The figures of the drawing among these the verbs apply to, each once</summary>
        static IFigure[] MovableAmong(IEnumerable<IFigure> figures)
        {
            return FigureParts.Wholes(figures)
                .Where(f => f.Drawing != null && IsMovable(f.Layer) && f.Drawing.Figures.Contains(f))
                .ToArray();
        }

        /// <summary>The figures of the bands as they are drawn, bottom up: by ZIndex, then the later in the list on top</summary>
        static List<IFigure> InDrawingOrder(Drawing drawing)
        {
            return drawing.Figures
                .Select((figure, index) => (Figure: figure, Index: index))
                .Where(f => IsMovable(f.Figure.Layer))
                .OrderBy(f => f.Figure.ZIndex)
                .ThenBy(f => f.Index)
                .Select(f => f.Figure)
                .ToList();
        }

        /// <summary>Whether a figure that is not among these is drawn over one of them (or under one) in its band</summary>
        static bool HasOtherBeyond(IEnumerable<IFigure> figures, bool over)
        {
            var moving = MovableAmong(figures).ToHashSet();
            if (moving.Count == 0)
            {
                return false;
            }

            var drawn = InDrawingOrder(moving.First().Drawing);
            if (!over)
            {
                drawn.Reverse();
            }

            // the bands where one of these was passed on the way
            var reached = new HashSet<ZOrder?>();
            foreach (var figure in drawn)
            {
                var band = Band(figure.Layer);
                if (moving.Contains(figure))
                {
                    reached.Add(band);
                }
                else if (reached.Contains(band))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>Whether any of these has a figure of its band over it that is not among them</summary>
        public static bool CanBringToFront(IEnumerable<IFigure> figures)
        {
            return HasOtherBeyond(figures, over: true);
        }

        /// <summary>Whether any of these is drawn over a figure of its band that is not among them</summary>
        public static bool CanSendToBack(IEnumerable<IFigure> figures)
        {
            return HasOtherBeyond(figures, over: false);
        }

        /// <summary>Puts these over every other figure of their band: one Z above the band's highest, one undo step</summary>
        public static void BringToFront(IEnumerable<IFigure> figures)
        {
            SetZ(figures, band => band.Max(f => f.Z) + 1);
        }

        /// <summary>Puts these under every other figure of their band: one Z below the band's lowest, one undo step</summary>
        public static void SendToBack(IEnumerable<IFigure> figures)
        {
            SetZ(figures, band => band.Min(f => f.Z) - 1);
        }

        /// <summary>Gives these, band by band, the Z the function picks from the figures of the band</summary>
        static void SetZ(IEnumerable<IFigure> figures, Func<IEnumerable<IFigure>, int> pickZ)
        {
            var moving = MovableAmong(figures);
            if (moving.Length == 0)
            {
                return;
            }

            var drawing = moving[0].Drawing;
            using (Transaction.Create(drawing.ActionManager, delayed: false))
            {
                foreach (var group in moving.GroupBy(f => Band(f.Layer)))
                {
                    int z = pickZ(drawing.Figures.Where(f => Band(f.Layer) == group.Key));
                    foreach (var figure in group)
                    {
                        drawing.ActionManager.SetProperty(figure, nameof(IFigure.Z), z);
                    }
                }
            }
        }
    }
}
