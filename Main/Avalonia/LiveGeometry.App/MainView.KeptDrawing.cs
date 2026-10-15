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

    /// <summary>How long the history must stay as it is before the drawing is kept</summary>
    static readonly TimeSpan KeepDelay = TimeSpan.FromSeconds(1);

    /// <summary>The user's drawing changed since it was last kept</summary>
    bool hasUnkeptChanges;

    // what is in the store now, as last written or read: the same text again is not written
    string storedDrawingText;
    string storedDrawingName;

    void InitializeKeptDrawing()
    {
        if (!KeepsOwnDrawing)
        {
            return;
        }

        keptDrawingText = SettingsStore.Current.Get(KeptDrawingKey);
        keptDrawingName = SettingsStore.Current.Get(KeptDrawingNameKey);
        storedDrawingText = keptDrawingText;
        storedDrawingName = keptDrawingName;

        SettingsStore.Leaving += () =>
        {
            // also with nothing changed in the history while the drawing is on screen: the
            // view (a zoom is no undo step); parked, nothing of it can have changed
            if (hasUnkeptChanges || EditsUserDrawing)
            {
                HandleExceptions(KeepOwnDrawing);
            }
        };
    }

    /// <summary>The drawings of the editor are kept after every undo step: hooked when the editor is made (<see cref="CreateEditor"/>)</summary>
    void KeepDrawingOfEditor()
    {
        if (!KeepsOwnDrawing)
        {
            return;
        }

        var control = DrawingHost.DrawingControl;
        control.DrawingAttach += drawing => drawing.ActionManager.CollectionChanged += ActionManager_KeepDrawing;
        control.DrawingDetach += drawing => drawing.ActionManager.CollectionChanged -= ActionManager_KeepDrawing;
        if (control.Drawing != null)
        {
            control.Drawing.ActionManager.CollectionChanged += ActionManager_KeepDrawing;
        }
    }

    bool HasKeptDrawing => keptDrawingText != null;

    /// <summary>Whether the editor shows the user's own drawing: not a drawing of the gallery, nor the blank page in front of a parked one</summary>
    bool EditsUserDrawing => drawingHost?.CurrentDrawing is { } drawing && drawing == UserDrawing;

    void ActionManager_KeepDrawing(object sender, EventArgs e)
    {
        // only the editor's drawing is listened to; edits of a drawing of the gallery are
        // dropped, and the drawing parked behind it doesn't change
        if (EditsUserDrawing)
        {
            KeepOwnDrawingSoon();
        }
    }

    /// <summary>
    /// After a pause in the changes (<see cref="KeepDelay"/>, from the last one): an undo
    /// step may raise several changes of the history, and a drag of a figure or of the view
    /// changes it at every move (its moves merge into one step). Kept at the next idle
    /// moment, the drawing was saved at every move, and in the browser the drag went a few
    /// frames a second. The save itself is one job of the UI thread (<see cref="KeepOwnDrawing"/>):
    /// the throttle's timer only posts it.
    /// </summary>
    void KeepOwnDrawingSoon()
    {
        if (!KeepsOwnDrawing)
        {
            return;
        }

        hasUnkeptChanges = true;
        Throttle.Schedule(
            this,
            view => Dispatcher.UIThread.Post(view.KeepChangedDrawing),
            KeepDelay,
            ThrottleOptions.PostponeUntilNoNewRequests);
    }

    void KeepChangedDrawing()
    {
        if (hasUnkeptChanges)
        {
            HandleExceptions(KeepOwnDrawing);
        }
    }

    /// <summary>
    /// The user's own drawing: the one parked while they look around the gallery (still
    /// <see cref="OwnDrawing"/> once they are back in it), else the editor's unless a
    /// drawing of the gallery is open there. (Not simply the editor's when no drawing of the
    /// gallery is open: back from one, the editor holds a blank page and the user's drawing
    /// is parked.)
    /// </summary>
    Drawing UserDrawing => OwnDrawing ?? (CurrentSample == null ? drawingHost?.CurrentDrawing : null);

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
            && drawing == drawingHost?.CurrentDrawing
            && (DrawingHost.DrawingControl.ConstructionInProgress || drawing.IsRecordingTransaction))
        {
            return;
        }

        bool isEmpty = drawing == null || !drawing.Figures.Any(figure => !(figure is CartesianGrid));
        StoreDrawing(isEmpty ? null : drawing.SaveAsText(), isEmpty ? null : OwnFileName);
        hasUnkeptChanges = false;
    }

    /// <summary>
    /// Into the store, unless it holds that already: the browser's local storage takes the
    /// whole text through JavaScript on every write, and the page hidden with nothing
    /// changed, or a drag that ends where it began, would write it all again
    /// </summary>
    void StoreDrawing(string text, string name)
    {
        if (text != storedDrawingText)
        {
            SettingsStore.Current.Set(KeptDrawingKey, text);
            storedDrawingText = text;
        }

        if (name != storedDrawingName)
        {
            SettingsStore.Current.Set(KeptDrawingNameKey, name);
            storedDrawingName = name;
        }
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
            StoreDrawing(text: null, name: null);
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
