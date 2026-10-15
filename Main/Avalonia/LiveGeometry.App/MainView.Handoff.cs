using System;
using System.Text;
using System.Text.Json;

namespace LiveGeometry;

/// <summary>
/// What a page without the editor hands over as it sends the user here (web/index.html,
/// goToEditor): New, Open with the file picked there, Edit of the drawing it showed (a shared
/// one, /s/&lt;id&gt;). The browser keeps it in the session's storage, which main.js takes it
/// out of before the app starts; done once, at startup, when the page is the editor's. Each
/// replaces the kept drawing, as New does.
/// </summary>
public partial class MainView
{
    /// <summary>The handoff as JSON ({ action, name, data: base64 }), or null; the browser head sets it</summary>
    public static Func<string> TakeHandoff { get; set; }

    /// <summary>Does what was handed over; false when nothing was, or the page is not the editor's</summary>
    bool DoHandoff(string path)
    {
        var text = TakeHandoff?.Invoke();
        if (text == null || !IsOwnDrawingPath(path))
        {
            return false;
        }

        using var json = JsonDocument.Parse(text);
        var root = json.RootElement;
        var data = root.TryGetProperty("data", out var dataElement) ? Convert.FromBase64String(dataElement.GetString()) : null;
        switch (root.GetProperty("action").GetString())
        {
            case "open":
                OpenDrawing(root.GetProperty("name").GetString(), data);
                break;
            case "edit":
                ShowSharedDrawing(data);
                break;
            default:
                ShowNewDrawing(push: false);
                break;
        }

        return true;
    }

    /// <summary>
    /// A shared drawing as the user's own, read as <see cref="OpenDrawing"/> reads an .lgf but
    /// without a file: it has no name, and Save asks for one
    /// </summary>
    void ShowSharedDrawing(byte[] bytes)
    {
        var text = Utilities.StripByteOrderMark(new UTF8Encoding().GetString(bytes));
        var xml = DrawingControl.ParseDrawing(text, out var problem);
        if (xml == null)
        {
            ShowNewDrawing(push: false);
            DrawingHost.ShowHint("The shared drawing could not be read: " + problem);
            return;
        }

        CurrentSample = null;
        ShowEditor();
        OwnDrawing = null;
        OwnFileName = null;
        OwnFile = null;
        DrawingHost.DrawingControl.LoadDrawing(xml, "");

        // a drawing with a caption (one of the gallery, shared) is laid out for this window,
        // once, as an opened file is
        var drawing = DrawingHost.CurrentDrawing;
        if (GalleryDrawing.HasCaption(drawing))
        {
            GalleryDrawing.Fit(drawing, GalleryDrawing.GetPlane(text));
        }

        ReplaceKeptDrawing();
        UpdateTour();
        Publish(OwnDrawingPath, OwnTitle, push: false);
    }
}
