using System;
using Avalonia;
using Avalonia.Controls;

namespace LiveGeometry;

/// <summary>
/// Rows of equal tiles that always fill the width: as many columns as fit at
/// <see cref="MinTileWidth"/>, then the tiles grow to share what is left over.
/// </summary>
public class TileGridPanel : Panel
{
    public double MinTileWidth { get; set; } = 210;

    public int MaxColumns { get; set; } = 6;

    public double Gap { get; set; } = 18;

    /// <summary>Height of a tile for a width of 1</summary>
    public double TileAspect { get; set; } = 0.8;

    int columns;
    Size tileSize;

    void Calculate(double availableWidth)
    {
        if (double.IsInfinity(availableWidth))
        {
            availableWidth = MinTileWidth * 4 + Gap * 3;
        }

        columns = (int)Math.Floor((availableWidth + Gap) / (MinTileWidth + Gap));
        columns = Math.Max(1, Math.Min(MaxColumns, columns));
        var width = Math.Floor((availableWidth - (columns - 1) * Gap) / columns);
        tileSize = new Size(width, Math.Round(width * TileAspect));
    }

    protected override Size MeasureOverride(Size availableSize)
    {
        Calculate(availableSize.Width);
        foreach (var child in Children)
        {
            child.Measure(tileSize);
        }

        int rows = (Children.Count + columns - 1) / columns;
        return new Size(
            columns * tileSize.Width + (columns - 1) * Gap,
            rows == 0 ? 0 : rows * tileSize.Height + (rows - 1) * Gap);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        Calculate(finalSize.Width);
        for (int i = 0; i < Children.Count; i++)
        {
            int row = i / columns;
            int column = i % columns;
            Children[i].Arrange(new Rect(
                column * (tileSize.Width + Gap),
                row * (tileSize.Height + Gap),
                tileSize.Width,
                tileSize.Height));
        }

        return finalSize;
    }
}
