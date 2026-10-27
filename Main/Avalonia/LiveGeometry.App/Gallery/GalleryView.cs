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
using Avalonia.Threading;
using AvaloniaEllipse = Avalonia.Controls.Shapes.Ellipse;

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

        // tiles in the first rows show emoji: their font is fetched now, and they wait for it
        // while the tiles without emoji load (NextTileToLoad)
        EmojiFont.EnsureLoaded().ContinueWith(
            _ => Dispatcher.UIThread.Post(ScheduleLoad),
            System.Threading.Tasks.TaskScheduler.Default);

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
        if (arrange)
        {
            var heading = new TextBlock()
            {
                Text = "Arrange the gallery: drag a tile into place. Every drop rewrites the order in GalleryCatalog.cs; rebuild to see it in the app.",
                FontSize = 16,
                FontWeight = FontWeight.SemiBold,
                Margin = HeadingMargin
            };
            heading.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Accent));
            content.Children.Add(heading);
        }
        else
        {
            content.Children.Add(CreateHeadingRow());
        }

        content.Children.Add(galleryTiles);
        if (arrange)
        {
            galleryTiles.AddHandler(PointerPressedEvent, Arrange_PointerPressed, RoutingStrategies.Tunnel);
            galleryTiles.PointerMoved += Arrange_PointerMoved;
            galleryTiles.PointerReleased += Arrange_PointerReleased;
        }

        // the part of the tiles in view, which the page scrolls and the window resizes
        galleryTiles.EffectiveViewportChanged += (s, e) =>
        {
            tilesInView = e.EffectiveViewport;
            ScheduleLoad();
        };

        // the tiles skip theme changes while the page is hidden behind the editor, and
        // nothing is loaded then
        this.GetObservable(IsVisibleProperty).Subscribe(new AnonymousObserver<bool>(visible =>
        {
            if (visible)
            {
                foreach (var picture in pictures)
                {
                    picture.RefreshThemeIfStale();
                }

                ScheduleLoad();
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

    #region Loading the drawings

    // The drawings of the tiles are loaded one at a time, when the UI thread has nothing
    // better to do, and only as they come into view: each load holds the UI thread (in the
    // browser a big drawing took most of a second), so a tile far down the page waits until
    // it is scrolled near, and nothing is loaded while the page is hidden behind the editor.

    Rect tilesInView;
    bool isLoadPosted;

    // Tiles loaded in one turn of the UI thread. Every turn ends with a layout and a frame:
    // fewer frames leave more of the time to the loading, and the tiles come in groups.
    const int TilesPerTurn = 8;

    // The tiles wait until the page has had a frame. In the browser the first batch was
    // posted by the first layout and ran in the same task as the first frame (Avalonia's
    // dispatcher runs what is due before it gives the page back, a short timer too), so the
    // browser could not paint until it was over: Avalonia's own splash, which it closed at
    // the first frame, stayed up a second longer, for the eight tiles behind it. The second
    // frame comes in a task of its own, after the first one is painted. (The page's splash now
    // stays until the tiles in view are in, SplashScreen, but the frame still lets the bar
    // under it move.)
    bool isOnScreen;

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        if (isOnScreen)
        {
            return;
        }

        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null)
        {
            isOnScreen = true;
            return;
        }

        topLevel.RequestAnimationFrame(_ => topLevel.RequestAnimationFrame(_ =>
        {
            isOnScreen = true;
            ScheduleLoad();
        }));
    }

    void ScheduleLoad()
    {
        if (isLoadPosted || !isOnScreen)
        {
            return;
        }

        isLoadPosted = true;
        Dispatcher.UIThread.Post(LoadNextTiles, DispatcherPriority.Background);
    }

    /// <summary>
    /// Loads the tiles nearest to the view that are not loaded yet: the ones in it first, from
    /// the top, then the ones up to half a screen below or above it, so that a scroll finds
    /// them there. A drawing with emoji waits for their font while the others load: loaded
    /// before it came, it showed its emoji seconds later. While the page's splash is up
    /// (<see cref="SplashScreen"/>), one tile a turn, so that the bar under the construction
    /// moves after each, and the splash comes down once the tiles in view are all loaded.
    /// </summary>
    void LoadNextTiles()
    {
        isLoadPosted = false;
        var tiles = TilesByDistance();
        double reach = tilesInView.Height / 2;
        bool isFontComing = !EmojiFont.EnsureLoaded().IsCompleted;
        bool IsDue(TileDistance tile) => tile.Picture.CanLoad && tile.Distance <= reach && !(isFontComing && tile.Picture.UsesEmoji);
        int perTurn = SplashScreen.IsUp ? 1 : TilesPerTurn;

        // a stable sort: the tiles at one distance stay in the grid's order
        foreach (var tile in tiles.OrderBy(tile => tile.Distance))
        {
            if (perTurn == 0)
            {
                break;
            }

            if (IsDue(tile))
            {
                tile.Picture.Load();
                perTurn--;
            }
        }

        if (SplashScreen.IsUp && tiles.Count > 0)
        {
            int inView = tiles.Count(tile => tile.Distance == 0);
            int loadedInView = tiles.Count(tile => tile.Distance == 0 && tile.Picture.HasLoaded);
            SplashScreen.ReportProgress(inView == 0 ? 1 : (double)loadedInView / inView);
            if (loadedInView == inView)
            {
                SplashScreen.Hide();
            }
        }

        if (tiles.Any(IsDue))
        {
            ScheduleLoad();
        }
    }

    struct TileDistance
    {
        public DrawingThumbnail Picture;

        /// <summary>0 for a tile in view, else how far it is above or below the view</summary>
        public double Distance;
    }

    /// <summary>
    /// The tiles in the grid's order, each with its distance from the part of the grid in
    /// view. From the grid's arithmetic (<see cref="TileGridPanel.GetTileRect"/>), not from
    /// each tile's place in the tree: this runs at every step of a scroll, and asking sixty
    /// tiles where they were took 30 ms on a phone. Empty while the page is hidden or the
    /// grid isn't laid out yet.
    /// </summary>
    List<TileDistance> TilesByDistance()
    {
        var result = new List<TileDistance>();
        if (!IsEffectivelyVisible || tilesInView.Height <= 0)
        {
            return result;
        }

        var tiles = galleryTiles.Children;
        for (int i = 0; i < tiles.Count; i++)
        {
            if (tiles[i] is GalleryTile tile && tile.Picture is DrawingThumbnail picture)
            {
                var rect = galleryTiles.GetTileRect(i);
                if (rect.Height == 0)
                {
                    result.Clear();
                    return result;
                }

                double distance = System.Math.Max(0, System.Math.Max(tilesInView.Top - rect.Bottom, rect.Top - tilesInView.Bottom));
                result.Add(new TileDistance { Picture = picture, Distance = distance });
            }
        }

        return result;
    }

    #endregion

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

        // a tile not loaded yet may have come into view
        ScheduleLoad();
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

    static readonly Thickness HeadingMargin = new Thickness(2, 30, 0, 12);

    /// <summary>A puzzle library of the site, a page of its own, as the start page links it</summary>
    /// <param name="Plate">The tint of its chip</param>
    /// <param name="Count">How many puzzles it holds</param>
    /// <param name="Picture">The miniature on its chip</param>
    public record Library(string Title, string Url, Color Plate, string Count, Func<Control> Picture);

    static Library[] Libraries()
    {
        const string catriona = "https://livegeometry.com/CatrionaAgg";
        // the paper of the /CatrionaAgg page
        var paper = Color.Parse("#F7F5F0");

        return
        [
            new Library("Catriona Agg's Puzzles", catriona, paper, "440", HexagonsPicture)
        ];
    }

    /// <summary>The library's page, in a new tab in the browser head and in the browser on the desktop</summary>
    void OpenLibrary(string url)
    {
        TopLevel.GetTopLevel(this)?.Launcher.LaunchUriAsync(new Uri(url));
    }

    /// <summary>The "Gallery" heading with a chip per puzzle library to its right</summary>
    Control CreateHeadingRow()
    {
        var row = new HeadingRow()
        {
            Margin = HeadingMargin
        };
        row.Children.Add(CreateHeading());
        row.Children.Add(CreateChips());
        return row;
    }

    /// <summary>
    /// The heading and, a gap to its right, the chips in a line - or, where the line is too
    /// narrow for all of them (a phone), the chips under the heading, one under another
    /// </summary>
    public class HeadingRow : Panel
    {
        const double Gap = 28;
        const double LineGap = 12;

        protected override Size MeasureOverride(Size availableSize)
        {
            var heading = Children[0];
            var chips = (StackPanel)Children[1];
            heading.Measure(Size.Infinity);
            double chipsWidth = 0;
            foreach (var chip in chips.Children)
            {
                chip.Measure(Size.Infinity);
                chipsWidth += chip.DesiredSize.Width;
            }

            bool oneLine = double.IsInfinity(availableSize.Width)
                || heading.DesiredSize.Width + Gap + chipsWidth <= availableSize.Width;
            chips.Orientation = oneLine ? Orientation.Horizontal : Orientation.Vertical;
            chips.Measure(Size.Infinity);
            double width = double.IsInfinity(availableSize.Width)
                ? heading.DesiredSize.Width + Gap + chips.DesiredSize.Width
                : availableSize.Width;
            double height = oneLine
                ? System.Math.Max(heading.DesiredSize.Height, chips.DesiredSize.Height)
                : heading.DesiredSize.Height + LineGap + chips.DesiredSize.Height;
            return new Size(width, height);
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            var heading = Children[0];
            var chips = (StackPanel)Children[1];
            var headingSize = heading.DesiredSize;
            var chipsSize = chips.DesiredSize;
            if (chips.Orientation == Orientation.Horizontal)
            {
                heading.Arrange(new Rect(0, (finalSize.Height - headingSize.Height) / 2, headingSize.Width, headingSize.Height));
                chips.Arrange(new Rect(headingSize.Width + Gap, (finalSize.Height - chipsSize.Height) / 2, chipsSize.Width, chipsSize.Height));
            }
            else
            {
                heading.Arrange(new Rect(0, 0, headingSize.Width, headingSize.Height));
                chips.Arrange(new Rect(0, headingSize.Height + LineGap, chipsSize.Width, chipsSize.Height));
            }

            return finalSize;
        }
    }

    static TextBlock CreateHeading()
    {
        var heading = new TextBlock()
        {
            Text = "Gallery",
            FontSize = 24,
            FontWeight = FontWeight.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        };
        heading.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
        return heading;
    }

    /// <summary>
    /// A pill per library on a pastel plate like the tiles: its miniature, its title and its
    /// count. In a line, or one under another (<see cref="HeadingRow"/> decides).
    /// </summary>
    Control CreateChips()
    {
        var panel = new StackPanel()
        {
            Orientation = Orientation.Horizontal
        };
        foreach (var library in Libraries())
        {
            var chip = new GalleryTile(library.Picture(), library.Title, library.Plate, () => OpenLibrary(library.Url))
            {
                IsChip = true,
                Badge = library.Count,
                HorizontalAlignment = HorizontalAlignment.Left,
                // the same above and below, so that a line of chips is centered on the
                // heading as the chip itself would be, and stacked chips stand apart
                Margin = new Thickness(0, 4, 10, 4)
            };
            panel.Children.Add(chip);
        }

        return panel;
    }

    // The fills of the hexagons miniature under the light and the dark theme: yellow, red,
    // blue, then the overlaps green (blue and yellow), purple (blue and red), orange (red and
    // yellow), and the hole inside all three. The light ones are the maintainer's own
    // (2026-10-27), with a hairline of ink on every edge; the dark ones are deep and vivid,
    // and tell the regions apart without ink.
    static readonly string[] hexagonLightFills = ["#F5F551", "#E97FA8", "#77D5FB", "#AFF73C", "#9975FF", "#FFBC7C", "#FFFFFF"];
    static readonly string[] hexagonDarkFills = ["#F2C000", "#E02A66", "#1B8FE6", "#5FC02A", "#5E34E8", "#F27D1E", "#2B2B2B"];
    static readonly IBrush hexagonInk = new SolidColorBrush(Color.Parse("#1E1E1E"));

    /// <summary>
    /// The three hexagons of Catriona Agg's "Three Regular Hexagons and a Semicircle" in
    /// miniature, laid out as the drawing is (the blue one's corner at 73° on the arc of the
    /// semicircle, which isn't drawn): the hexagons, their overlaps and the hole in the middle,
    /// every edge in a hairline of ink under the light theme, the height of a chip's picture,
    /// a pixel of room around for the ink.
    /// </summary>
    static Control HexagonsPicture()
    {
        const double radius = 4;
        const double sqrt3 = 1.7320508075688772;
        const double s3 = sqrt3 / 2;
        const double size = 26;
        Point Polar(double degrees)
        {
            return new Point(radius * System.Math.Cos(degrees * System.Math.PI / 180), radius * System.Math.Sin(degrees * System.Math.PI / 180));
        }

        Point Turned(Point point, double degrees)
        {
            double cos = System.Math.Cos(degrees * System.Math.PI / 180);
            double sin = System.Math.Sin(degrees * System.Math.PI / 180);
            return new Point(point.X * cos - point.Y * sin, point.X * sin + point.Y * cos);
        }

        var y = Polar(73);
        var r = Polar(133);
        var p = new Point(y.X - y.Y / sqrt3, 0);
        double ys = y.Y / sqrt3;
        double rs = r.Y / sqrt3;
        Point[] yellowCorners =
        [
            p,
            p + new Point(ys, 0),
            p + new Point(1.5 * ys, s3 * ys),
            p + new Point(ys, 2 * s3 * ys),
            p + new Point(0, 2 * s3 * ys),
            p + new Point(-0.5 * ys, s3 * ys)
        ];
        Point[] redCorners =
        [
            p,
            p + new Point(-rs, 0),
            p + new Point(-1.5 * rs, s3 * rs),
            r,
            p + new Point(0, 2 * s3 * rs),
            p + new Point(0.5 * rs, s3 * rs)
        ];
        var center = new Point((y.X + r.X) / 2, (y.Y + r.Y) / 2);
        var arm = y - center;
        Point[] blueCorners = Enumerable.Range(0, 6).Select(i => center + Turned(arm, 60 * i)).ToArray();

        // the drawing's y goes up; the canvas's goes down, in pixels, from a pixel above the top
        var corners = yellowCorners.Concat(redCorners).Concat(blueCorners).ToList();
        double left = corners.Min(corner => corner.X);
        double top = corners.Max(corner => corner.Y);
        double scale = (size - 2) / (top - corners.Min(corner => corner.Y));
        Geometry Hexagon(Point[] hexagon)
        {
            return new PolylineGeometry(hexagon.Select(corner => new Point((corner.X - left) * scale + 1, (top - corner.Y) * scale + 1)).ToList(), isFilled: true);
        }

        var yellow = Hexagon(yellowCorners);
        var red = Hexagon(redCorners);
        var blue = Hexagon(blueCorners);
        var blueAndRed = new CombinedGeometry(GeometryCombineMode.Intersect, blue, red);
        var canvas = new Canvas()
        {
            Width = (corners.Max(corner => corner.X) - left) * scale + 2,
            Height = size,
            Margin = new Thickness(10, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        Path[] fills =
        [
            new Path() { Data = yellow },
            new Path() { Data = red },
            new Path() { Data = blue },
            new Path() { Data = new CombinedGeometry(GeometryCombineMode.Intersect, blue, yellow) },
            new Path() { Data = blueAndRed },
            new Path() { Data = new CombinedGeometry(GeometryCombineMode.Intersect, red, yellow) },
            new Path() { Data = new CombinedGeometry(GeometryCombineMode.Intersect, blueAndRed, yellow) }
        ];
        canvas.Children.AddRange(fills);
        Path[] edges = [Inked(yellow), Inked(red), Inked(blue)];
        canvas.Children.AddRange(edges);

        // the fills follow the theme on screen; the ink is the light theme's alone
        canvas.ObserveTheme(nameof(AppTheme.Page), page =>
        {
            bool dark = AppTheme.Current == AppTheme.Dark;
            string[] colors = dark ? hexagonDarkFills : hexagonLightFills;
            for (int i = 0; i < fills.Length; i++)
            {
                fills[i].Fill = new SolidColorBrush(Color.Parse(colors[i]));
            }

            foreach (var edge in edges)
            {
                edge.IsVisible = !dark;
            }
        });
        return canvas;
    }

    /// <summary>A hairline of ink along the geometry: a device pixel on a 200% screen</summary>
    static Path Inked(Geometry geometry)
    {
        return new Path()
        {
            Data = geometry,
            Stroke = hexagonInk,
            StrokeThickness = 0.5,
            StrokeJoin = PenLineJoin.Round
        };
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
        var disc = new AvaloniaEllipse() { Width = 84, Height = 84 };
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
