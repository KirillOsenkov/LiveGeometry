using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Reactive;
using DynamicGeometry;
using Ellipse = Avalonia.Controls.Shapes.Ellipse;

namespace LiveGeometry;

/// <summary>
/// The start page: New, Open, My Drawing (while there is one to go back to) and the drawings
/// of the gallery as tiles. Has no ribbon and no canvas of its own - <see cref="MainView"/>
/// shows either this or the editor.
/// </summary>
public class GalleryView : DockPanel
{
    public event Action NewDrawingRequested = delegate { };
    public event Action OpenDrawingRequested = delegate { };
    public event Action ContinueDrawingRequested = delegate { };
    public event Action<GalleryItem> ItemRequested = delegate { };

    static readonly Color newDrawingPlate = Color.Parse("#E6EFFB");
    static readonly Color openDrawingPlate = Color.Parse("#E4F4E6");
    static readonly Color continueDrawingPlate = Color.Parse("#FBF1DC");

    readonly StackPanel startTiles = new StackPanel()
    {
        Orientation = Orientation.Horizontal,
        Spacing = StartTileSpacing,
        VerticalAlignment = VerticalAlignment.Top
    };
    readonly TileGridPanel galleryTiles = new TileGridPanel();
    readonly List<DrawingThumbnail> pictures = new List<DrawingThumbnail>();
    readonly GalleryTile continueTile;

    /// <param name="arrange">The arrange mode: see <see cref="Arrange_PointerPressed"/></param>
    public GalleryView(Control corner, bool arrange = false)
    {
        this.BindTheme(BackgroundProperty, nameof(AppTheme.Page));

        startTiles.Children.Add(CreateStartTile(PlusPicture(), "New", newDrawingPlate, () => NewDrawingRequested()));
        startTiles.Children.Add(CreateStartTile(FolderPicture(), "Open", openDrawingPlate, () => OpenDrawingRequested()));
        continueTile = CreateStartTile(PencilPicture(), "My Drawing", continueDrawingPlate, () => ContinueDrawingRequested());
        continueTile.IsVisible = false;
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
            pictures.Add(picture);
            var tile = new GalleryTile(
                picture,
                item.Title,
                Pastels.At(i),
                () =>
                {
                    picture.IsAnimated = false;
                    ItemRequested(item);
                });
            tile.Tag = item;
            tile.PointerEntered += (s, e) => picture.IsAnimated = true;
            tile.PointerExited += (s, e) => picture.IsAnimated = false;

            // a drawing with paper of its own shows it on its tile instead of the pastel, as
            // the theme on screen resolves it (a Dark override of the paper counts)
            picture.DrawingLoaded += drawing =>
            {
                if (drawing.OwnBackground != null)
                {
                    tile.SetPlate(drawing.Background);
                    tilesWithPaper.Add(tile);
                }
            };
            AppTheme.CurrentChanged += () =>
            {
                if (picture.Drawing?.OwnBackground != null)
                {
                    tile.SetPlate(picture.Drawing.Background);
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
        var heading = new TextBlock()
        {
            Text = arrange
                ? "Arrange the gallery: drag a tile into place. Every drop rewrites the order in GalleryCatalog.cs; rebuild to see it in the app."
                : "Gallery",
            FontSize = arrange ? 16 : 24,
            FontWeight = FontWeight.SemiBold,
            Margin = new Thickness(2, 30, 0, 12)
        };
        heading.BindTheme(TextBlock.ForegroundProperty, arrange ? nameof(AppTheme.Accent) : nameof(AppTheme.Text));
        content.Children.Add(heading);
        content.Children.Add(galleryTiles);
        if (arrange)
        {
            galleryTiles.AddHandler(PointerPressedEvent, Arrange_PointerPressed, RoutingStrategies.Tunnel);
            galleryTiles.PointerMoved += Arrange_PointerMoved;
            galleryTiles.PointerReleased += Arrange_PointerReleased;
        }

        // the tiles skip theme changes while the page is hidden behind the editor
        this.GetObservable(IsVisibleProperty).Subscribe(new AnonymousObserver<bool>(visible =>
        {
            if (visible)
            {
                foreach (var picture in pictures)
                {
                    picture.RefreshThemeIfStale();
                }
            }
        }));

        var page = new Panel();
        page.Children.Add(new ScrollViewer()
        {
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = content
        });

        // the theme button and the build stamp in the corner, as on the editor's toolbar
        if (corner != null)
        {
            corner.HorizontalAlignment = HorizontalAlignment.Right;
            corner.VerticalAlignment = VerticalAlignment.Top;
            corner.Margin = new Thickness(0, 6, 22, 0);
            page.Children.Add(corner);
        }

        Children.Add(page);
    }

    // Three of them fit across the narrowest phone (360 wide, less the margins of the page),
    // where the brand gives up its place; "My Drawing" is as long as a caption gets
    const double StartTileSize = 92;
    const double StartTileSpacing = 10;
    const double StartPictureSize = 44;

    static GalleryTile CreateStartTile(Control picture, string text, Color plate, Action action)
    {
        return new GalleryTile(picture, text, plate, action)
        {
            Width = StartTileSize,
            Height = StartTileSize,
            IsCompact = true
        };
    }

    #region Arrange mode

    // A way to reorder the gallery by hand ("LiveGeometry.Desktop.exe --arrange"): a tile is
    // dragged and takes the place of whatever tile the pointer is over, the grid reordering
    // as it goes; a drop writes the order into the catalog's source. The pointer is captured
    // on press, so the tiles never see a click and nothing opens.

    GalleryTile dragged;
    readonly HashSet<GalleryTile> tilesWithPaper = new HashSet<GalleryTile>();

    void Arrange_PointerPressed(object sender, PointerPressedEventArgs e)
    {
        if (!e.GetCurrentPoint(galleryTiles).Properties.IsLeftButtonPressed)
        {
            return;
        }

        dragged = TileAt(e.GetPosition(galleryTiles));
        if (dragged == null)
        {
            return;
        }

        dragged.Opacity = 0.5;
        e.Pointer.Capture(galleryTiles);
        e.Handled = true;
    }

    void Arrange_PointerMoved(object sender, PointerEventArgs e)
    {
        if (dragged == null)
        {
            return;
        }

        var target = TileAt(e.GetPosition(galleryTiles));
        if (target == null || target == dragged)
        {
            return;
        }

        var tiles = galleryTiles.Children;
        tiles.Move(tiles.IndexOf(dragged), tiles.IndexOf(target));

        // the pastels go by position, as they will after the rebuild
        for (int i = 0; i < tiles.Count; i++)
        {
            if (tiles[i] is GalleryTile tile && !tilesWithPaper.Contains(tile))
            {
                tile.SetPlate(Pastels.At(i));
            }
        }
    }

    void Arrange_PointerReleased(object sender, PointerReleasedEventArgs e)
    {
        if (dragged == null)
        {
            return;
        }

        dragged.Opacity = 1;
        dragged = null;
        e.Pointer.Capture(null);
        var order = galleryTiles.Children.OfType<GalleryTile>().Select(tile => (GalleryItem)tile.Tag);
        Console.WriteLine("gallery order written to " + GalleryCatalog.SaveOrder(order));
    }

    GalleryTile TileAt(Point position)
    {
        return galleryTiles.Children.OfType<GalleryTile>().FirstOrDefault(tile => tile.Bounds.Contains(position));
    }

    #endregion

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
        var name = new TextBlock()
        {
            Text = "Live Geometry",
            FontSize = 38,
            FontWeight = FontWeight.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        };
        name.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.TextMuted));
        brand.Children.Add(name);

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
        var disc = new Ellipse() { Width = 84, Height = 84 };
        disc.ObserveTheme(nameof(AppTheme.Accent), color => disc.Fill = Lit(color));
        return Picture(
            disc,
            new Path()
            {
                Data = Geometry.Parse("M42,22 V62 M22,42 H62"),
                Stroke = Brushes.White,
                StrokeThickness = 7,
                StrokeLineCap = PenLineCap.Round
            });
    }

    /// <summary>The folder of the toolbar's Open, in its colors</summary>
    static Control FolderPicture()
    {
        var outline = new SolidColorBrush(Color.FromRgb(0xA8, 0x7B, 0x05));
        return Picture(
            new Path()
            {
                Data = Geometry.Parse("M10,18 H33 L40,27 H70 V66 H10 Z"),
                Fill = new SolidColorBrush(Color.FromRgb(0xF2, 0xB6, 0x32)),
                Stroke = outline,
                StrokeThickness = 2.5,
                StrokeJoin = PenLineJoin.Round
            },
            new Path()
            {
                Data = Geometry.Parse("M10,66 L20,38 H79 L70,66 Z"),
                Fill = Lit(Color.FromRgb(0xFA, 0xD5, 0x65)),
                Stroke = outline,
                StrokeThickness = 2.5,
                StrokeJoin = PenLineJoin.Round
            });
    }

    /// <summary>
    /// The color as if lit from the upper left, as the tool icons' shapes are: a diagonal
    /// gradient from a lighter tint of it there to the color itself at the lower right.
    /// </summary>
    /// <param name="lightAt">Where along the diagonal of the shape's box the tint is, 0 to 1</param>
    /// <param name="fullAt">Where the color itself is: a shape that lies across the diagonal
    /// (the pencil) spans only the middle of it, and wants the blend in there</param>
    static LinearGradientBrush Lit(Color color, double lightAt = 0, double fullAt = 1)
    {
        return new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Lighter(color, 0.6), lightAt),
                new GradientStop(color, fullAt)
            }
        };
    }

    /// <summary>The color mixed with white, by the given part of the way</summary>
    static Color Lighter(Color color, double part)
    {
        return Color.FromArgb(
            color.A,
            (byte)(color.R + (255 - color.R) * part),
            (byte)(color.G + (255 - color.G) * part),
            (byte)(color.B + (255 - color.B) * part));
    }

    static Control PencilPicture()
    {
        // the pencil runs from the lower left to the upper right, across the light: on the
        // diagonal of its box it takes up only the part between 0.39 and 0.61
        var wood = Lit(Color.FromRgb(0xF2, 0xB6, 0x32), lightAt: 0.39, fullAt: 0.61);
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
            MaxWidth = StartPictureSize,
            MaxHeight = StartPictureSize,
            Margin = new Thickness(6, 8, 6, 0),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
    }
}
