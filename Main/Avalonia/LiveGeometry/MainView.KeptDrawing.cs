using System;
using System.Linq;
using Avalonia.Threading;
using DynamicGeometry;
using Drawing = DynamicGeometry.Drawing;

namespace LiveGeometry;

/// <summary>
/// The user's own drawing kept between visits, in the settings store (the browser's local
/// storage): a reload, or coming back to the site another day, finds it where it was left -
/// at /drawing, or behind the gallery's My Drawing. It is kept after every undo step (and
/// when the page is hidden, for the view), as the .lgf text, with the name of the file it
/// came from; a drawing with nothing in it is no drawing to keep. Read at startup but loaded
/// only when the user goes to it, like a drawing parked behind the gallery.
/// </summary>
public partial class MainView
{
    /// <summary>
    /// Whether the user's own drawing is kept between runs: the browser head turns it on.
    /// The desktop has files for that (and its settings file is a line per key).
    /// </summary>
    public static bool KeepsOwnDrawing { get; set; }

    const string KeptDrawingKey = "Drawing";
    const string KeptDrawingNameKey = "DrawingName";

    /// <summary>
    /// The kept drawing as stored, until it is loaded or replaced (a new drawing, a file
    /// opened): meanwhile it is the user's own drawing, and nothing is kept over it
    /// </summary>
    string keptDrawingText;

    string keptDrawingName;

    bool keepScheduled;

    void InitializeKeptDrawing()
    {
        if (!KeepsOwnDrawing)
        {
            return;
        }

        keptDrawingText = SettingsStore.Current.Get(KeptDrawingKey);
        keptDrawingName = SettingsStore.Current.Get(KeptDrawingNameKey);

        var control = DrawingHost.DrawingControl;
        control.DrawingAttach += drawing => drawing.ActionManager.CollectionChanged += ActionManager_KeepDrawing;
        control.DrawingDetach += drawing => drawing.ActionManager.CollectionChanged -= ActionManager_KeepDrawing;
        if (control.Drawing != null)
        {
            control.Drawing.ActionManager.CollectionChanged += ActionManager_KeepDrawing;
        }

        SettingsStore.Leaving += () => HandleExceptions(KeepOwnDrawing);
    }

    bool HasKeptDrawing => keptDrawingText != null;

    void ActionManager_KeepDrawing(object sender, EventArgs e)
    {
        KeepOwnDrawingSoon();
    }

    /// <summary>Once the dispatcher is idle: an undo step may raise several changes of the history</summary>
    void KeepOwnDrawingSoon()
    {
        if (!KeepsOwnDrawing || keepScheduled)
        {
            return;
        }

        keepScheduled = true;
        Dispatcher.UIThread.Post(() =>
        {
            keepScheduled = false;
            HandleExceptions(KeepOwnDrawing);
        }, DispatcherPriority.Background);
    }

    /// <summary>
    /// The user's own drawing: the one parked while they look around the gallery (still
    /// <see cref="OwnDrawing"/> once they are back in it), else the editor's unless a
    /// drawing of the gallery is open there. (Not simply the editor's when no drawing of the
    /// gallery is open: back from one, the editor holds a blank page and the user's drawing
    /// is parked.)
    /// </summary>
    Drawing UserDrawing => OwnDrawing ?? (CurrentSample == null ? DrawingHost.CurrentDrawing : null);

    void KeepOwnDrawing()
    {
        if (!KeepsOwnDrawing || HasKeptDrawing)
        {
            return;
        }

        var drawing = UserDrawing;

        // a construction under way is not in the drawing yet (its point following the
        // cursor is a figure while it lasts): kept when it is done, which is an undo step
        if (drawing != null
            && drawing == DrawingHost.CurrentDrawing
            && (DrawingHost.DrawingControl.ConstructionInProgress || drawing.IsRecordingTransaction))
        {
            return;
        }

        bool isEmpty = drawing == null || !drawing.Figures.Any(figure => !(figure is CartesianGrid));
        SettingsStore.Current.Set(KeptDrawingKey, isEmpty ? null : drawing.SaveAsText());
        SettingsStore.Current.Set(KeptDrawingNameKey, isEmpty ? null : OwnFileName);
    }

    /// <summary>The user's drawing is another one now: the kept one is replaced by it</summary>
    void ReplaceKeptDrawing()
    {
        keptDrawingText = null;
        keptDrawingName = null;
        KeepOwnDrawingSoon();
    }

    /// <summary>
    /// The kept drawing into the editor, as the user's own (<see cref="ShowOwnDrawing"/> when
    /// nothing is parked); false when its text is no drawing, which is dropped then
    /// </summary>
    bool ShowKeptDrawing(bool push)
    {
        var text = keptDrawingText;
        var name = keptDrawingName;
        keptDrawingText = null;
        keptDrawingName = null;
        var xml = DrawingControl.ParseDrawing(text, out _);
        if (xml == null)
        {
            SettingsStore.Current.Set(KeptDrawingKey, null);
            SettingsStore.Current.Set(KeptDrawingNameKey, null);
            return false;
        }

        CurrentSample = null;
        ShowEditor();
        OwnDrawing = null;
        OwnFileName = name;
        OwnFile = null;
        DrawingHost.DrawingControl.LoadDrawing(xml, name ?? "");
        UpdateTour();
        Publish(OwnDrawingPath, OwnTitle, push);
        return true;
    }
}
