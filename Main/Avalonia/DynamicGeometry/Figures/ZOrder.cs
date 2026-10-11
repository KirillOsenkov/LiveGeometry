using System.Collections.Generic;
using System.Linq;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    /// <summary>
    /// The layers figures are drawn in, bottom up. A figure's <see cref="IFigure.Layer"/> is
    /// its kind's; within the band from <see cref="Polygons"/> to <see cref="Vectors"/> its
    /// <see cref="IFigure.Z"/> (0 unless Bring to front or Send to back changed it) comes
    /// first and the layer second (<see cref="ZOrders.Encode"/>), so a polygon brought to front
    /// covers a circle, and nothing in the band ever covers a point. Within one value the
    /// later figure is on top.
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
        Points = 60,
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

        /// <summary>All the Z values of the band, with their layers</summary>
        const int Span = Stride * (2 * Range + 1);

        /// <summary>Whether a figure of the layer can be brought to front and sent to back</summary>
        public static bool IsMovable(ZOrder layer)
        {
            return layer >= ZOrder.Polygons && layer <= ZOrder.Vectors;
        }

        /// <summary>
        /// The ZIndex of the layer at the Z: the layers below the band as they are, the band
        /// above them ordered by Z and then by layer, the layers above the band over all of it.
        /// </summary>
        public static int Encode(ZOrder layer, int z)
        {
            if (layer < ZOrder.Polygons)
            {
                return (int)layer;
            }

            if (layer > ZOrder.Vectors)
            {
                return Stride + Span + (int)layer;
            }

            return Stride + (System.Math.Clamp(z, -Range, Range) + Range) * Stride + (int)layer;
        }

        /// <summary>The ZIndex of a figure of the layer whose Z was never changed</summary>
        public static int Default(ZOrder layer)
        {
            return Encode(layer, 0);
        }

        /// <summary>The figures of the drawing the verbs order among themselves (a part counts as its figure)</summary>
        public static IEnumerable<IFigure> Movable(Drawing drawing)
        {
            return drawing.Figures.Where(f => IsMovable(f.Layer));
        }

        /// <summary>The figures of the drawing among these the verbs apply to, each once</summary>
        static IFigure[] MovableAmong(IEnumerable<IFigure> figures)
        {
            return FigureParts.Wholes(figures)
                .Where(f => f.Drawing != null && IsMovable(f.Layer) && f.Drawing.Figures.Contains(f))
                .ToArray();
        }

        /// <summary>Whether the first is drawn over the second: by ZIndex, then the later in the list</summary>
        static bool IsOver(IFigure figure, IFigure other)
        {
            if (figure.ZIndex != other.ZIndex)
            {
                return figure.ZIndex > other.ZIndex;
            }

            var figures = figure.Drawing.Figures;
            return figures.IndexOf(figure) > figures.IndexOf(other);
        }

        /// <summary>Whether any of these has a figure over it that is not among them</summary>
        public static bool CanBringToFront(IEnumerable<IFigure> figures)
        {
            var moving = MovableAmong(figures);
            return moving.Length > 0
                && Movable(moving[0].Drawing).Any(other => !moving.Contains(other) && moving.Any(f => IsOver(other, f)));
        }

        /// <summary>Whether any of these is drawn over a figure that is not among them</summary>
        public static bool CanSendToBack(IEnumerable<IFigure> figures)
        {
            var moving = MovableAmong(figures);
            return moving.Length > 0
                && Movable(moving[0].Drawing).Any(other => !moving.Contains(other) && moving.Any(f => IsOver(f, other)));
        }

        /// <summary>Puts these over every other figure of the band: one Z above the highest, one undo step</summary>
        public static void BringToFront(IEnumerable<IFigure> figures)
        {
            var moving = MovableAmong(figures);
            if (moving.Length == 0)
            {
                return;
            }

            var drawing = moving[0].Drawing;
            int z = Movable(drawing).Max(f => f.Z) + 1;
            SetZ(drawing, moving, z);
        }

        /// <summary>Puts these under every other figure of the band: one Z below the lowest, one undo step</summary>
        public static void SendToBack(IEnumerable<IFigure> figures)
        {
            var moving = MovableAmong(figures);
            if (moving.Length == 0)
            {
                return;
            }

            var drawing = moving[0].Drawing;
            int z = Movable(drawing).Min(f => f.Z) - 1;
            SetZ(drawing, moving, z);
        }

        static void SetZ(Drawing drawing, IFigure[] figures, int z)
        {
            using (Transaction.Create(drawing.ActionManager, delayed: false))
            {
                foreach (var figure in figures)
                {
                    drawing.ActionManager.SetProperty(figure, nameof(IFigure.Z), z);
                }
            }
        }
    }
}
