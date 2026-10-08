using System;
using System.IO;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using DynamicGeometry;

namespace LiveGeometry;

/// <summary>
/// The Export button of the toolbar: the canvas as it is on screen at the moment - the same
/// view, the same size - saved as a picture or put on the clipboard. And Save as, the
/// drawing under another name, which is here since Save stopped asking for one.
/// </summary>
public partial class MainView
{
    static readonly FilePickerFileType PngFileType = new("PNG image")
    {
        Patterns = new[] { "*.png" },
        MimeTypes = new[] { "image/png" }
    };

    static readonly FilePickerFileType SvgFileType = new("SVG image")
    {
        Patterns = new[] { "*.svg" },
        MimeTypes = new[] { "image/svg+xml" }
    };

    static readonly FilePickerFileType HtmlFileType = new("HTML page")
    {
        Patterns = new[] { "*.html" },
        MimeTypes = new[] { "text/html" }
    };

    MainToolbarButton ExportButton;

    /// <summary>
    /// What the clipboard holds is asked for when someone pastes, which can be long after
    /// the copy: the picture lives until the next one is copied
    /// </summary>
    Bitmap copiedImage;

    void ShowExportMenu()
    {
        var menu = new MenuFlyout() { Placement = PlacementMode.BottomEdgeAlignedLeft };

        void Add(string header, Action action)
        {
            var item = new MenuItem() { Header = header };
            item.Click += (s, e) => action();
            menu.Items.Add(item);
        }

        Add("Save as .lgf", SaveDrawingAs);
        Add("Save as .png", SaveAsPng);
        Add("Save as .svg", SaveAsSvg);
        Add("Save as .html", SaveAsHtml);
        menu.Items.Add(new Separator());
        Add("Copy image", CopyImage);
        Add("Copy embed code", CopyEmbedCode);

        // the button stays down while its menu is open
        menu.Closed += (s, e) => ExportButton.IsChecked = false;
        ExportButton.IsChecked = true;
        menu.ShowAt(ExportButton);
    }

    /// <summary>In the pixels of the screen: at 200% scaling the picture is twice the canvas's layout size</summary>
    byte[] RenderPng()
    {
        var canvas = DrawingHost.DrawingControl;
        return ViewportImage.ToPng(canvas, TopLevel.GetTopLevel(canvas).RenderScaling);
    }

    void SaveAsPng()
    {
        SaveImage(
            "Save as PNG",
            "png",
            PngFileType,
            () => Task.FromResult(RenderPng()));
    }

    void SaveAsSvg()
    {
        SaveImage(
            "Save as SVG",
            "svg",
            SvgFileType,
            () => ViewportImage.ToSvg(DrawingHost.DrawingControl));
    }

    async void SaveImage(
        string title,
        string extension,
        FilePickerFileType fileType,
        Func<Task<byte[]>> render)
    {
        try
        {
            // first, so that the picture is of the moment of the click and not of whatever
            // has moved while the dialog was open
            var bytes = await render();
            var name = CurrentSample != null ? CurrentSample.FileName : OwnFileName;
            var topLevel = TopLevel.GetTopLevel(this);
            var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = title,
                SuggestedFileName = Path.ChangeExtension(name ?? "drawing", extension),
                DefaultExtension = extension,
                FileTypeChoices = new[] { fileType }
            });

            if (file == null)
            {
                return;
            }

            if (await TryWriteFile(file, bytes))
            {
                DrawingHost.ShowHint("Saved " + file.Name);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    /// <summary>
    /// The drawing as Save writes it. A construction under way is not in the drawing yet:
    /// its point following the cursor and its preview are figures like any other while it
    /// lasts, and went into the file ("TempPoint", half a polygon). It is put away, as
    /// Escape does.
    /// </summary>
    string DrawingTextToSave()
    {
        var drawing = DrawingHost.CurrentDrawing;
        if (DrawingHost.DrawingControl.ConstructionInProgress || drawing.IsRecordingTransaction)
        {
            drawing.Behavior?.Restart();
        }

        return drawing.SaveAsText();
    }

    /// <summary>The drawing's title for a page: the gallery's, or the file's name</summary>
    string ExportTitle
    {
        get
        {
            return CurrentSample != null ? CurrentSample.Title : Path.GetFileNameWithoutExtension(OwnFileName ?? "drawing");
        }
    }

    /// <summary>
    /// A web page of its own, the player written into it (PlayerEmbed.CreatePage): one file
    /// that plays from anywhere, a disk included
    /// </summary>
    void SaveAsHtml()
    {
        SaveImage(
            "Save as HTML page",
            "html",
            HtmlFileType,
            async () =>
            {
                var page = PlayerEmbed.CreatePage(ExportTitle, DrawingTextToSave(), await PlayerEmbed.OpenPlayerScript());
                return new System.Text.UTF8Encoding(false).GetBytes(page);
            });
    }

    /// <summary>
    /// The embed snippet on the clipboard (PlayerEmbed.CreateSnippet): the drawing inside
    /// an element, played by the script on livegeometry.com; to paste into any HTML page.
    /// (GitHub's READMEs, issues and wikis strip scripts: there, a picture linking to a
    /// page that holds the snippet is the way.)
    /// </summary>
    void CopyEmbedCode()
    {
        try
        {
            // as tall as the drawing is on screen now, in the page's pixels
            int height = (int)System.Math.Max(240, System.Math.Round(DrawingHost.DrawingControl.Bounds.Height));
            Clipboard.SetText(PlayerEmbed.CreateSnippet(DrawingTextToSave(), height));
            DrawingHost.ShowHint("Embed code copied: paste it into any HTML page.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    async void CopyImage()
    {
        try
        {
            // decoded again into a plain bitmap, which every clipboard can read the pixels of
            var image = new Bitmap(new MemoryStream(RenderPng()));
            await TopLevel.GetTopLevel(this).Clipboard.SetBitmapAsync(image);
            copiedImage?.Dispose();
            copiedImage = image;
            DrawingHost.ShowHint("Image copied");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }
}
