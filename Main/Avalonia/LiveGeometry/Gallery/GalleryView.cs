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

    // around the color wheel in steps of the golden angle: neighbors never look alike
    const double hueStep = 137.5;
    const double newDrawingHue = 215;
    const double continueDrawingHue = 40;

    static readonly IBrush accent = new SolidColorBrush(Color.FromRgb(0x2F, 0x7B, 0xD6));
    static readonly IBrush secondaryText = new SolidColorBrush(Color.FromRgb(0x5B, 0x64, 0x72));

    readonly TileGridPanel startTiles = new TileGridPanel() { TileAspect = 0.56 };
    readonly TileGridPanel galleryTiles = new TileGridPanel();
    readonly GalleryTile continueTile;

    public GalleryView(Control buildStamp)
    {
        Background = Brushes.White;

        var header = CreateHeader(buildStamp);
        SetDock(header, Dock.Top);
        Children.Add(header);

        startTiles.Children.Add(new GalleryTile(PlusPicture(), "New Drawing", newDrawingHue, () => NewDrawingRequested()));
        continueTile = new GalleryTile(PencilPicture(), "Back to My Drawing", continueDrawingHue, () => ContinueDrawingRequested())
        {
            IsVisible = false
        };
        startTiles.Children.Add(continueTile);

        for (int i = 0; i < GalleryCatalog.Items.Count; i++)
        {
            var item = GalleryCatalog.Items[i];
            var picture = new DrawingThumbnail(item) { Margin = new Thickness(8, 8, 8, 4) };
            var tile = new GalleryTile(
                picture,
                item.Title,
                i * hueStep,
                () =>
                {
                    picture.IsAnimated = false;
                    ItemRequested(item);
                });
            tile.PointerEntered += (s, e) => picture.IsAnimated = true;
            tile.PointerExited += (s, e) => picture.IsAnimated = false;
            galleryTiles.Children.Add(tile);
        }

        var content = new StackPanel()
        {
            MaxWidth = 1500,
            Margin = new Thickness(32, 24, 32, 40)
        };
        content.Children.Add(startTiles);
        content.Children.Add(new TextBlock()
        {
            Text = "Gallery",
            FontSize = 24,
            FontWeight = FontWeight.SemiBold,
            Foreground = RibbonTheme.Text,
            Margin = new Thickness(2, 30, 0, 2)
        });
        content.Children.Add(new TextBlock()
        {
            Text = "Every picture is alive. Open one, drag the yellow points and watch what changes - and what never does.",
            FontSize = 15,
            Foreground = secondaryText,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(2, 0, 0, 16)
        });
        content.Children.Add(galleryTiles);

        Children.Add(new ScrollViewer()
        {
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = content
        });
    }

    /// <summary>
    /// Whether there is a drawing of the user's own to go back to (the editor keeps it while
    /// the gallery is showing)
    /// </summary>
    public bool CanContinueDrawing
    {
        get => continueTile.IsVisible;
        set => continueTile.IsVisible = value;
    }

    // the same band as the editor's toolbar, so that switching between the two is quiet
    static Control CreateHeader(Control buildStamp)
    {
        var header = new DockPanel()
        {
            Background = RibbonTheme.HeaderRowBackground,
            Height = 44
        };

        if (buildStamp != null)
        {
            SetDock(buildStamp, Dock.Right);
            header.Children.Add(buildStamp);
        }

        header.Children.Add(new TextBlock()
        {
            Text = "Live Geometry",
            FontSize = 18,
            FontWeight = FontWeight.SemiBold,
            Foreground = RibbonTheme.Text,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(34, 0, 0, 0)
        });

        var line = new Border()
        {
            Height = 1,
            Background = RibbonTheme.TabLine,
            VerticalAlignment = VerticalAlignment.Bottom
        };

        var panel = new Panel();
        panel.Children.Add(header);
        panel.Children.Add(line);
        return panel;
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
            MaxWidth = 84,
            MaxHeight = 84,
            Margin = new Thickness(16),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
    }
}
