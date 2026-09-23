using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Shapes;
using Avalonia.Layout;
using Avalonia.Media;
using DynamicGeometry;
using Ellipse = Avalonia.Controls.Shapes.Ellipse;

namespace LiveGeometry;

/// <summary>
/// The start page: "New Drawing" and the drawings of the gallery as tiles. Has no ribbon and no
/// canvas of its own - <see cref="MainView"/> shows either this or the editor.
/// </summary>
public class GalleryView : DockPanel
{
    public event Action NewDrawingRequested = delegate { };
    public event Action ContinueDrawingRequested = delegate { };
    public event Action<GalleryItem> ItemRequested = delegate { };

    static readonly Color newDrawingPlate = Color.Parse("#E6EFFB");
    static readonly Color continueDrawingPlate = Color.Parse("#FBF1DC");

    static readonly IBrush accent = new SolidColorBrush(Color.FromRgb(0x2F, 0x7B, 0xD6));
    static readonly IBrush brandText = new SolidColorBrush(Color.FromRgb(0x6B, 0x74, 0x82));

    readonly StackPanel startTiles = new StackPanel()
    {
        Orientation = Orientation.Horizontal,
        Spacing = 18,
        VerticalAlignment = VerticalAlignment.Top
    };
    readonly TileGridPanel galleryTiles = new TileGridPanel();
    readonly GalleryTile continueTile;

    public GalleryView(Control buildStamp)
    {
        Background = Brushes.White;

        startTiles.Children.Add(new GalleryTile(PlusPicture(), "New Drawing", newDrawingPlate, () => NewDrawingRequested())
        {
            Width = StartTileWidth,
            Height = StartTileHeight
        });
        continueTile = new GalleryTile(PencilPicture(), "My Drawing", continueDrawingPlate, () => ContinueDrawingRequested())
        {
            Width = StartTileWidth,
            Height = StartTileHeight,
            IsVisible = false
        };
        startTiles.Children.Add(continueTile);

        // the start row: the tiles at the left, the name of the app in the room to their right
        var startRow = new Grid()
        {
            ColumnDefinitions = new ColumnDefinitions("Auto,*")
        };
        startRow.Children.Add(startTiles);
        var brand = CreateBrand();
        Grid.SetColumn(brand, 1);
        startRow.Children.Add(brand);

        for (int i = 0; i < GalleryCatalog.Items.Count; i++)
        {
            var item = GalleryCatalog.Items[i];
            var picture = new DrawingThumbnail(item) { Margin = new Thickness(8, 8, 8, 4) };
            var tile = new GalleryTile(
                picture,
                item.Title,
                Pastels.At(i),
                () =>
                {
                    picture.IsAnimated = false;
                    ItemRequested(item);
                });
            tile.PointerEntered += (s, e) => picture.IsAnimated = true;
            tile.PointerExited += (s, e) => picture.IsAnimated = false;

            // a drawing with paper of its own shows it on its tile instead of the pastel
            picture.Loaded += drawing =>
            {
                if (!DynamicGeometry.Drawing.IsWhite(drawing.Background))
                {
                    tile.SetPlate(drawing.Background);
                }
            };
            galleryTiles.Children.Add(tile);
        }

        var content = new StackPanel()
        {
            MaxWidth = 1500,
            Margin = new Thickness(32, 24, 32, 40)
        };
        content.Children.Add(startRow);
        content.Children.Add(new TextBlock()
        {
            Text = "Gallery",
            FontSize = 24,
            FontWeight = FontWeight.SemiBold,
            Foreground = RibbonTheme.Text,
            Margin = new Thickness(2, 30, 0, 12)
        });
        content.Children.Add(galleryTiles);

        var page = new Panel();
        page.Children.Add(new ScrollViewer()
        {
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = content
        });

        // the build stamp in the corner, as on the editor's toolbar
        if (buildStamp != null)
        {
            buildStamp.HorizontalAlignment = HorizontalAlignment.Right;
            buildStamp.VerticalAlignment = VerticalAlignment.Top;
            buildStamp.Margin = new Thickness(0, 6, 22, 0);
            page.Children.Add(buildStamp);
        }

        Children.Add(page);
    }

    // square, about the height of the brand next to them
    const double StartTileWidth = 120;
    const double StartTileHeight = 120;

    /// <summary>
    /// Whether there is a drawing of the user's own to go back to (the editor keeps it while
    /// the gallery is showing)
    /// </summary>
    public bool CanContinueDrawing
    {
        get => continueTile.IsVisible;
        set => continueTile.IsVisible = value;
    }

    /// <summary>The icon and the name, floating on a soft shadow</summary>
    static Control CreateBrand()
    {
        var brand = new StackPanel()
        {
            Orientation = Orientation.Horizontal,
            Spacing = 20
        };
        brand.Children.Add(AppIcon.Create(size: 68));
        brand.Children.Add(new TextBlock()
        {
            Text = "Live Geometry",
            FontSize = 38,
            FontWeight = FontWeight.SemiBold,
            Foreground = brandText,
            VerticalAlignment = VerticalAlignment.Center
        });

        // shrinks on a narrow window instead of running off the edge; the margin keeps it
        // off the tiles at its left
        var fitted = new Viewbox()
        {
            Child = brand,
            StretchDirection = StretchDirection.DownOnly,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };

        // The shadow is cut off at the bounds of the element that carries it, so that element
        // is a frame with room for the blur all around; the negative margin gives the room
        // back, the frame takes no more space in the layout than the brand.
        const double blurRoom = 40;
        return new Border()
        {
            Child = fitted,
            Padding = new Thickness(blurRoom),
            Margin = new Thickness(36 - blurRoom, -blurRoom, 12 - blurRoom, -blurRoom),
            Effect = new DropShadowEffect()
            {
                Color = Colors.Black,
                Opacity = 0.32,
                BlurRadius = 26,
                OffsetX = 0,
                OffsetY = 6
            }
        };
    }

    static Control PlusPicture()
    {
        return Picture(
            new Ellipse() { Width = 84, Height = 84, Fill = accent },
            new Path()
            {
                Data = Geometry.Parse("M42,22 V62 M22,42 H62"),
                Stroke = Brushes.White,
                StrokeThickness = 7,
                StrokeLineCap = PenLineCap.Round
            });
    }

    static Control PencilPicture()
    {
        var wood = new SolidColorBrush(Color.FromRgb(0xF2, 0xB6, 0x32));
        var outline = new SolidColorBrush(Color.FromRgb(0x8A, 0x63, 0x05));
        return Picture(
            new Path()
            {
                Data = Geometry.Parse("M16,68 L20,52 L58,14 L70,26 L32,64 Z"),
                Fill = wood,
                Stroke = outline,
                StrokeThickness = 2.5,
                StrokeJoin = PenLineJoin.Round
            },
            new Path()
            {
                Data = Geometry.Parse("M20,52 L32,64 M50,22 L62,34"),
                Stroke = outline,
                StrokeThickness = 2.5,
                StrokeLineCap = PenLineCap.Round
            });
    }

    static Control Picture(params Control[] shapes)
    {
        var canvas = new Canvas() { Width = 84, Height = 84 };
        canvas.Children.AddRange(shapes);
        return new Viewbox()
        {
            Child = canvas,
            Stretch = Stretch.Uniform,
            MaxWidth = 56,
            MaxHeight = 56,
            Margin = new Thickness(8, 10, 8, 2),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
    }
}
