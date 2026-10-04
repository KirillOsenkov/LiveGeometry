using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using DynamicGeometry;
using Drawing = DynamicGeometry.Drawing;

namespace LiveGeometry;

/// <summary>
/// The shared Live Geometry main view, used by both the Browser and Desktop heads.
/// Ported from Main/WPFClient/MainWindow.cs.
/// </summary>
public partial class MainView : UserControl
{
    DockPanel LayoutRoot = new DockPanel();
    DrawingHost drawingHost;

    /// <summary>
    /// The editor: the canvas, the ribbon with every tool, the side panel, the toolbar. Built
    /// the first time anything asks for it (<see cref="CreateEditor"/>), not at startup: the
    /// gallery, the start page, showed only once all of it had been made, and a visitor who
    /// only looks at the gallery never needs it.
    /// </summary>
    DrawingHost DrawingHost => drawingHost ?? CreateEditor();

    /// <summary>Whether the editor is there yet: what only reaches into it (a key, a press, a resize) asks first, so as not to build it</summary>
    bool IsEditorBuilt => drawingHost != null;

    Behavior[] Behaviors = Array.Empty<Behavior>();

    // The browser's picker (File System Access API) takes a type only as a MIME type with its
    // extensions; Avalonia drops a type without MimeTypes there, so the save dialog offered
    // no .lgf at all.
    static readonly FilePickerFileType LgfFileType = new("Live Geometry drawing")
    {
        Patterns = new[] { "*.lgf" },
        MimeTypes = new[] { "application/xml" }
    };

    static readonly FilePickerFileType AnyDrawingFileType = new("All drawings")
    {
        Patterns = new[] { "*.lgf", "*.dgf", "*.ggb" },
        MimeTypes = new[] { "application/xml", "text/plain", "application/vnd.geogebra.file" }
    };

    static readonly FilePickerFileType DgfFileType = new("DG 1.x drawing")
    {
        Patterns = new[] { "*.dgf" },
        MimeTypes = new[] { "text/plain" }
    };

    static readonly FilePickerFileType GgbFileType = new("GeoGebra worksheet")
    {
        Patterns = new[] { "*.ggb" },
        MimeTypes = new[] { "application/vnd.geogebra.file" }
    };

    public MainView()
    {
        InitializeComponent();
        InitializeKeptDrawing();

        // The geometry library surfaces errors through the WPF-style MessageBox shim: in the
        // status bar, which only the editor has
        MessageBox.Handler = text =>
        {
            if (IsEditorBuilt)
            {
                DrawingHost.ShowHint(text);
            }
            else
            {
                Console.WriteLine(text);
            }
        };
        AppDomain.CurrentDomain.FirstChanceException += CurrentDomain_FirstChanceException;

        AddHandler(KeyDownEvent, MainView_KeyDown, RoutingStrategies.Tunnel);
        AddHandler(KeyUpEvent, MainView_KeyUp, RoutingStrategies.Tunnel);
        AddHandler(PointerPressedEvent, MainView_PointerPressed, RoutingStrategies.Tunnel, handledEventsToo: true);

        // a phone turned on its side is a different screen
        SizeChanged += (s, e) => UpdateRibbon();

        AttachedToVisualTree += (s, e) =>
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel != null)
            {
                DynamicGeometry.Clipboard.SystemClipboard = topLevel.Clipboard;
                topLevel.AddHandler(KeyDownEvent, TopLevel_KeyDown, RoutingStrategies.Tunnel);
            }

            Focus();

            // once the canvas has a size, or the drawing would be laid out in a 0x0 viewport
            Avalonia.Threading.Dispatcher.UIThread.Post(OpenStartupFile, Avalonia.Threading.DispatcherPriority.Loaded);
        };
    }

    /// <summary>The editor, the first time it is needed (see <see cref="DrawingHost"/>)</summary>
    DrawingHost CreateEditor()
    {
        drawingHost = new DrawingHost();
        AddBehaviors();

        // the toolbar on top, the drawing host filling the rest
        CreateToolbar();
        LayoutRoot.Children.Add(drawingHost);
        InitializeCommands();
        KeepDrawingOfEditor();
        drawingHost.UnhandledException += (s, e) => ReportException(e.Exception);

        // Give the drawing canvas keyboard focus so behaviors receive Escape/Delete/etc.
        drawingHost.DrawingControl.Focusable = true;
        drawingHost.DrawingControl.PointerPressed += (s, e) => drawingHost.DrawingControl.Focus();

        // the tour group and the ribbon as the page at hand wants them
        UpdateTour();
        return drawingHost;
    }

    private void AddBehaviors()
    {
        var behaviors = Behavior.LoadBehaviors(typeof(Dragger).Assembly);
        Behaviors = behaviors.ToArray();
        Behavior.Default = behaviors.First(b => b is Dragger);
        foreach (var behavior in behaviors)
        {
            DrawingHost.AddToolButton(behavior);
        }

        // the user's own tools, after the library's: on Misc, behind Define figure
        var storedTools = new StoredTools();
        ToolStorage.Instance = storedTools;
        storedTools.Load();
    }

    // A link to the repository at the far right of the toolbar; its tooltip is the build (git
    // commit) on screen - e.g. whether a fresh deployment has arrived yet. A faint Octocat
    // rather than the raw commit hash, which meant nothing to the kids the app is for.
    public const string RepositoryUrl = "https://github.com/KirillOsenkov/LiveGeometry";

    const double OctocatSize = 16;
    const double OctocatOpacity = 0.4;

    // the GitHub mark (the "mark-github" octicon, on a 16 x 16 grid)
    const string OctocatPath = "M8,0 C3.58,0 0,3.58 0,8 c0,3.54 2.29,6.53 5.47,7.59 c0.4,0.07 0.55,-0.17 0.55,-0.38 c0,-0.19 -0.01,-0.82 -0.01,-1.49 c-2.01,0.37 -2.53,-0.49 -2.69,-0.94 c-0.09,-0.23 -0.48,-0.94 -0.82,-1.13 c-0.28,-0.15 -0.68,-0.52 -0.01,-0.53 c0.63,-0.01 1.08,0.58 1.23,0.82 c0.72,1.21 1.87,0.87 2.33,0.66 c0.07,-0.52 0.28,-0.87 0.51,-1.07 c-1.78,-0.2 -3.64,-0.89 -3.64,-3.95 c0,-0.87 0.31,-1.59 0.82,-2.15 c-0.08,-0.2 -0.36,-1.02 0.08,-2.12 c0,0 0.67,-0.21 2.2,0.82 c0.64,-0.18 1.32,-0.27 2,-0.27 c0.68,0 1.36,0.09 2,0.27 c1.53,-1.04 2.2,-0.82 2.2,-0.82 c0.44,1.1 0.16,1.92 0.08,2.12 c0.51,0.56 0.82,1.27 0.82,2.15 c0,3.07 -1.87,3.75 -3.65,3.95 c0.29,0.25 0.54,0.73 0.54,1.48 c0,1.07 -0.01,1.93 -0.01,2.2 c0,0.21 0.15,0.46 0.55,0.38 A8.013,8.013 0 0 0 16,8 c0,-4.42 -3.58,-8 -8,-8 z";

    static Control CreateBuildStamp()
    {
        var mark = new Avalonia.Controls.Shapes.Path()
        {
            Data = Geometry.Parse(OctocatPath),
            Stretch = Stretch.Uniform,
            Width = OctocatSize,
            Height = OctocatSize
        };
        mark.BindTheme(Avalonia.Controls.Shapes.Shape.FillProperty, nameof(AppTheme.Text));
        var octocat = CreateCornerButton(mark, () => TopLevel.GetTopLevel(mark)?.Launcher.LaunchUriAsync(new Uri(RepositoryUrl)));
        ToolTip.SetTip(octocat, BuildVersion.Full + "\n" + RepositoryUrl);
        return octocat;
    }

    /// <summary>
    /// The sun/moon beside the Octocat, on both pages: the other of light and dark, at a click
    /// (the settings page has the full choice, System included).
    /// </summary>
    static Control CreateThemeButton()
    {
        var picture = new Viewbox()
        {
            Width = OctocatSize,
            Height = OctocatSize
        };
        var button = CreateCornerButton(picture, AppSettings.Instance.ToggleDarkTheme);

        void Update()
        {
            bool isDark = AppTheme.Current == AppTheme.Dark;
            picture.Child = isDark ? MainToolbarIcons.Sun() : MainToolbarIcons.Moon();
            ToolTip.SetTip(button, isDark ? "Light theme" : "Dark theme");
        }

        Update();
        AppTheme.CurrentChanged += Update;
        return button;
    }

    /// <summary>A small faint picture that comes to life under the pointer; the corner of both pages holds a couple</summary>
    static Border CreateCornerButton(Control picture, Action action)
    {
        var button = new CornerButton()
        {
            Background = Brushes.Transparent, // hit-testable around the picture too
            Cursor = new Cursor(StandardCursorType.Hand),
            Opacity = OctocatOpacity,
            Margin = new Avalonia.Thickness(6, 0, 6, 0),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            Child = picture
        };
        button.PointerEntered += (s, e) => button.Opacity = 1;
        button.PointerExited += (s, e) => button.Opacity = OctocatOpacity;
        button.PointerReleased += (s, e) =>
        {
            if (e.InitialPressMouseButton == MouseButton.Left)
            {
                action();
            }
        };
        return button;
    }

    /// <summary>A type of its own, so that a press on it is a press on a button and not on empty chrome (<see cref="IsLayoutOnly"/>)</summary>
    class CornerButton : Border
    {
        protected override Type StyleKeyOverride => typeof(Border);
    }

    /// <summary>The theme button and the build stamp, for the corner of either page</summary>
    static Control CreateCorner()
    {
        var corner = new StackPanel()
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Margin = new Avalonia.Thickness(0, 0, 4, 0)
        };
        corner.Children.Add(CreateThemeButton());
        corner.Children.Add(CreateBuildStamp());
        return corner;
    }

    /// <summary>The gear: the settings page in the side panel, or away again</summary>
    void ToggleSettings()
    {
        var settings = AppSettings.Instance;
        DrawingHost.ShowProperties(DrawingHost.PropertyGrid.Selection == settings ? null : settings);
    }

    private void InitializeComponent()
    {
        Focusable = true;
        Console.WriteLine("Live Geometry " + BuildVersion.Full);

        // Two pages, one showing: the gallery (the start page) and the editor. Started with a
        // file (or a batch job), or at the address of a drawing (a shared link to one of the
        // gallery, or /drawing), the editor is up from the first frame and the gallery, with
        // its tiles, isn't even built until the Gallery button is pressed; started at the
        // gallery, the editor isn't built until a drawing is opened.
        pages.Children.Add(LayoutRoot);
        Content = pages;
        bool startsInEditor = StartupFile != null || CheckFolder != null || ModernizeFolder != null || RewriteFolder != null || RecaptionFolder != null || SpaceLabelsFolder != null || IsDrawingPath(AddressBar.Current.Path);
        LayoutRoot.IsVisible = startsInEditor;
        if (startsInEditor)
        {
            CreateEditor();
        }
        else
        {
            EnsureGallery();
        }
    }

    void CreateToolbar()
    {
        // No menu: the few document commands are a toolbar, everything else is the keyboard
        // (see MainView_KeyUp and HandlePlainKey), the mouse wheel and the context menu.
        var toolbar = Toolbar = new MainToolbar();
        toolbar.AddAtRight(CreateCorner());
        AppSettings.Instance.ShowRequested += page => HandleExceptions(() => DrawingHost.ShowProperties(page));
        AppSettings.Instance.DrawingBackgroundRequested += () => HandleExceptions(DrawingHost.ShowDrawingProperties);
        ToolboxButton = toolbar.AddButton(
            AppIcon.Create(BrandIconSize),
            "Tools",
            RibbonShortcut,
            ToggleRibbon,
            iconSize: BrandIconSize,
            inset: MainToolbarButton.DefaultInset - (BrandIconSize - MainToolbarButton.IconSize) / 2);
        toolbar.SetTabButton(ToolboxButton);
        toolbar.AddButton(MainToolbarIcons.Gallery(), "Gallery", shortcut: null, () => HandleExceptions(() => ShowGallery(push: true)));
        toolbar.AddSeparator();
        toolbar.AddButton(MainToolbarIcons.New(), "New", "Ctrl+N", NewDrawing);
        toolbar.AddButton(MainToolbarIcons.Open(), "Open", "Ctrl+O", OpenDrawingFromFile);
        toolbar.AddButton(MainToolbarIcons.Save(), "Save", "Ctrl+S", SaveDrawing);
        ExportButton = toolbar.AddButton(MainToolbarIcons.Export(), "Export", shortcut: null, ShowExportMenu);
        toolbar.AddSeparator();
        toolbar.AddButton(MainToolbarIcons.Undo(), "Ctrl+Z", DrawingHost.DrawingControl.CommandUndo);
        toolbar.AddButton(MainToolbarIcons.Redo(), "Ctrl+Y", DrawingHost.DrawingControl.CommandRedo);
        toolbar.AddSeparator();
        toolbar.AddButton(MainToolbarIcons.Settings(), "Settings", shortcut: null, ToggleSettings);

        // while a drawing of the gallery is open: previous / next through the gallery and its
        // title, after the buttons (or in the middle of the room they leave, as the toolbar's
        // IsGroupLeftAligned says), and bigger than the document
        // buttons; while a drawing from a file is open, the file's name alone
        TourGroup = toolbar.BeginCenteredGroup();
        TourPrevious = toolbar.AddButton(
            MainToolbarIcons.Previous(),
            "Previous drawing",
            "Page Up",
            () => HandleExceptions(() => ShowNeighborSample(-1)),
            iconSize: TourIconSize,
            inset: TourArrowInset);
        TourPosition = toolbar.AddText(FontWeight.Normal, minWidth: 44, fontSize: TourFontSize);
        TourNext = toolbar.AddButton(
            MainToolbarIcons.Next(),
            "Next drawing",
            "Page Down",
            () => HandleExceptions(() => ShowNeighborSample(1)),
            iconSize: TourIconSize,
            inset: TourArrowInset);
        TourTitle = toolbar.AddTrailingText(FontWeight.SemiBold, fontSize: TourFontSize);
        toolbar.EndGroup();
        TourGroup.IsVisible = false;

        LayoutRoot.Children.Add(toolbar);
        DockPanel.SetDock(toolbar, Dock.Top);
    }

    #region Ribbon

    // The ribbon can be folded away behind the first button of the toolbar, to leave a small
    // screen to the drawing. It starts folded when a drawing of the gallery is opened on a small
    // screen (one is there to look, and the tabs wouldn't fit anyway) and open everywhere else;
    // once the user has pressed the button, their choice holds for the rest of the session.

    /// <summary>The editor's, made with it (<see cref="CreateToolbar"/>)</summary>
    MainToolbar Toolbar;
    const string RibbonShortcut = "Ctrl+F1";

    /// <summary>
    /// The app's own mark, in the top left corner where a logo goes, doubling as the toggle. A
    /// little bigger than the document icons, in a button of the same height.
    /// </summary>
    MainToolbarButton ToolboxButton;
    const double BrandIconSize = 28;

    /// <summary>Null until the user toggles the ribbon themselves</summary>
    bool? ribbonChoice;

    const double SmallScreenWidth = 700;
    const double SmallScreenHeight = 500;

    bool IsSmallScreen => Bounds.Width > 0 && (Bounds.Width < SmallScreenWidth || Bounds.Height < SmallScreenHeight);

    void ToggleRibbon()
    {
        ribbonChoice = !DrawingHost.Ribbon.IsVisible;
        UpdateRibbon();
    }

    void UpdateRibbon()
    {
        if (!IsEditorBuilt)
        {
            return;
        }

        bool visible = ribbonChoice ?? !(CurrentSample != null && IsSmallScreen);
        DrawingHost.Ribbon.IsVisible = visible;
        Toolbar.IsTabOpen = visible;
        ToolTip.SetTip(ToolboxButton, (visible ? "Hide the tools" : "Show the tools") + " (" + RibbonShortcut + ")");
    }

    #endregion

    void InitializeCommands()
    {
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleGrid, first: true);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleLabelNewPoints);
        DrawingHost.AddToolbarButton(DrawingHost.CommandTogglePointByCoordinates);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleFigureExplorer);

        // Not on the ribbon for now: Ortho, Polar, Snap to grid / point / center
        // (DrawingHost.CommandToggle*). Their settings still work (Shift = snap to grid).
    }

    public void HandleExceptions(Action code)
    {
        try
        {
            code();
        }
        catch (Exception e)
        {
            MessageBox.Show(e.Message);
        }
    }

    #region Exceptions

    // Every exception thrown anywhere, at the moment it is thrown (as in Helix and the
    // structured log viewer): the message goes to the status bar and the whole text to the
    // side panel, so that a problem the code swallows is still seen. Reported on the UI
    // thread; while a report is on its way, exceptions it may throw itself are dropped, and
    // the same exception thrown again and again (on every mouse move, say) only refreshes
    // the status bar.

    bool reportingException;
    string lastExceptionText;

    void CurrentDomain_FirstChanceException(object sender, System.Runtime.ExceptionServices.FirstChanceExceptionEventArgs e)
    {
        var exception = e.Exception;
        if (reportingException || exception == null || IsBenign(exception))
        {
            return;
        }

        // the exception's own trace is filled in as it unwinds, and in the browser it can
        // stay empty; here, at the throw, the stack still has the throw site on it
        var stackTrace = string.IsNullOrEmpty(exception.StackTrace) ? Environment.StackTrace : null;
        reportingException = true;
        try
        {
            Dispatcher.UIThread.Post(() =>
            {
                try
                {
                    ReportException(exception, stackTrace);
                }
                finally
                {
                    reportingException = false;
                }
            });
        }
        catch
        {
            reportingException = false;
        }
    }

    static bool IsBenign(Exception exception)
    {
        return exception is OperationCanceledException
            || exception is AggregateException
            || exception is System.Reflection.TargetInvocationException
            // an element of a GeoGebra worksheet left out, which the status says
            || exception is GeoGebraReader.LeftOutException
            // the browser's file picker throws on cancel (Avalonia catches it and answers null)
            || exception.GetType().Name == "JSException" && exception.Message.StartsWith("AbortError", StringComparison.Ordinal)
            || IsReadOnlyFile(exception)
            // a file that can't be written now, said in words (TryWriteFile)
            || writingFile && (exception is IOException || exception is UnauthorizedAccessException);
    }

    /// <summary>The message in the status bar, the whole text in the side panel</summary>
    /// <param name="stackTrace">The stack at the throw, for an exception whose own trace is empty</param>
    public void ReportException(Exception exception, string stackTrace = null)
    {
        var report = new ExceptionReport(exception, stackTrace);
        var text = report.Details;
        bool isNew = text != lastExceptionText;
        lastExceptionText = text;
        if (isNew)
        {
            Console.WriteLine("LiveGeometry error: " + text);
        }

        // the status bar and the side panel are the editor's; on the gallery before it, the
        // console has it
        if (!IsEditorBuilt)
        {
            return;
        }

        DrawingHost.ShowHint("Error: " + exception.Message);
        if (isNew && DrawingHost.CurrentDrawing != null)
        {
            DrawingHost.ShowProperties(report);
        }
    }

    #endregion

    void NewDrawing() => HandleExceptions(() => ShowNewDrawing(push: true));

    #region Pages

    // Where the app is: the gallery, a drawing of the gallery ("the tour": previous / next in
    // the toolbar, edits are thrown away without asking) or a drawing of the user's own.
    // Every change goes through one of the Show methods, which also keep the address bar
    // current; `push` is false when the address bar is where the change came from.

    const string GalleryPath = "/";
    const string OwnDrawingPath = "/drawing";
    const string AppTitle = "Live Geometry";

    readonly Panel pages = new Panel();

    /// <summary>Null until first shown</summary>
    GalleryView Gallery;

    bool IsGalleryShowing => Gallery != null && Gallery.IsVisible;

    void EnsureGallery()
    {
        if (Gallery != null)
        {
            return;
        }

        Gallery = new GalleryView(CreateCorner(), ArrangeGallery);
        Gallery.NewDrawingRequested += () => HandleExceptions(() => ShowNewDrawing(push: true));
        Gallery.OpenDrawingRequested += OpenDrawingFromFile;
        Gallery.ContinueDrawingRequested += () => HandleExceptions(() => ShowOwnDrawing(push: true));
        Gallery.ItemRequested += item => HandleExceptions(() => ShowSample(item, push: true));
        pages.Children.Add(Gallery);
    }

    const double TourIconSize = 32;
    const double TourFontSize = 17;
    const double TourArrowInset = 1; // the chevrons have room enough inside their own icon

    Panel TourGroup;
    MainToolbarButton TourPrevious;
    TextBlock TourPosition;
    MainToolbarButton TourNext;
    TextBlock TourTitle;

    /// <summary>The drawing of the gallery that is open, null for a drawing of the user's own</summary>
    GalleryItem CurrentSample;

    /// <summary>
    /// The user's own drawing, kept (with its undo history) while they look around the gallery
    /// </summary>
    Drawing OwnDrawing;

    /// <summary>
    /// The name of the file the user's own drawing came from or was last saved to (with its
    /// extension), null for a new drawing
    /// </summary>
    string OwnFileName;

    /// <summary>
    /// The .lgf file the user's own drawing was read from or last saved to: Save writes it
    /// again without asking. Null for a new drawing and for one that came from another
    /// format (GeoGebra, DG), which Save has to ask a name for.
    /// </summary>
    IStorageFile OwnFile;

    /// <summary>The page title of the user's own drawing: the file's name, when it has one</summary>
    string OwnTitle => OwnFileName != null ? OwnFileName + " - " + AppTitle : AppTitle;

    static bool IsOwnDrawingPath(string path)
    {
        return path != null && string.Equals(path.TrimEnd('/'), OwnDrawingPath, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Whether the path shows the editor rather than the gallery</summary>
    static bool IsDrawingPath(string path)
    {
        return GalleryCatalog.FindByPath(path) != null || IsOwnDrawingPath(path);
    }

    void Navigate(string path, bool push)
    {
        var item = GalleryCatalog.FindByPath(path);
        if (item != null)
        {
            ShowSample(item, push);
        }
        else if (IsOwnDrawingPath(path))
        {
            ShowOwnDrawing(push);
        }
        else
        {
            ShowGallery(push);
            if (path != GalleryPath && path != GalleryCatalog.PathPrefix.TrimEnd('/'))
            {
                AddressBar.Current.Replace(GalleryPath, AppTitle);
            }
        }
    }

    void Publish(string path, string title, bool push)
    {
        if (push)
        {
            AddressBar.Current.Push(path, title);
        }
        else
        {
            AddressBar.Current.Replace(path, title);
        }
    }

    void ShowGallery(bool push)
    {
        var drawing = drawingHost?.CurrentDrawing;
        if (CurrentSample != null)
        {
            // nothing of the user's to keep; and the editor shouldn't show it when it comes back
            CurrentSample = null;
            DrawingHost.Clear();
        }
        else if (drawing != null && drawing.Figures.Any(figure => !(figure is CartesianGrid)))
        {
            OwnDrawing = drawing;
        }

        EnsureGallery();
        Gallery.CanContinueDrawing = OwnDrawing != null || HasKeptDrawing;
        Gallery.IsVisible = true;
        LayoutRoot.IsVisible = false;
        UpdateTour();
        Publish(GalleryPath, AppTitle, push);
    }

    void ShowEditor()
    {
        if (Gallery != null)
        {
            Gallery.IsVisible = false;
        }

        LayoutRoot.IsVisible = true;

        // the drawing skipped theme changes while the editor was hidden behind the gallery
        DrawingHost.CurrentDrawing?.RefreshThemeIfStale();

        // the canvas must know its size before a drawing is fitted into it, and the ribbon
        // (folded or not, which the callers have decided by setting CurrentSample) is part of it
        UpdateRibbon();
        UpdateLayout();
        DrawingHost.DrawingControl.Focus();
    }

    void ShowNewDrawing(bool push)
    {
        CurrentSample = null;
        ShowEditor();
        OwnDrawing = null;
        OwnFileName = null;
        OwnFile = null;
        DrawingHost.Clear();
        ReplaceKeptDrawing();
        UpdateTour();
        Publish(OwnDrawingPath, OwnTitle, push);
    }

    /// <summary>
    /// Back to the drawing the user left for the gallery, or kept from their last visit; a
    /// new one if there is none
    /// </summary>
    void ShowOwnDrawing(bool push)
    {
        if (OwnDrawing == null && HasKeptDrawing && ShowKeptDrawing(push))
        {
            return;
        }

        if (OwnDrawing == null)
        {
            ShowNewDrawing(push);
            return;
        }

        CurrentSample = null;
        ShowEditor();
        var control = DrawingHost.DrawingControl;
        if (control.Drawing != OwnDrawing)
        {
            // the same order as DrawingControl.Clear: on the canvas first (which brings the
            // drawing's own paper along), then the current one
            OwnDrawing.Canvas = control;
            control.Drawing = OwnDrawing;
            OwnDrawing.Recalculate();
        }

        UpdateTour();
        Publish(OwnDrawingPath, OwnTitle, push);
    }

    void ShowSample(GalleryItem item, bool push)
    {
        var control = DrawingHost.DrawingControl;
        if (OwnDrawing == null && CurrentSample == null && control.Drawing != null && control.Drawing.Figures.Any(figure => !(figure is CartesianGrid)))
        {
            OwnDrawing = control.Drawing;
        }

        CurrentSample = item;
        ShowEditor();
        control.LoadDrawing(item.LoadText(), item.FileName);

        // a reader drags the figure, not the text: on a phone a thumb on the caption scrolls
        control.Drawing.FixLabels();
        var drawing = control.Drawing;
        drawing.FitToWindow = () => GalleryDrawing.Fit(drawing, item.Plane, item.StackedFigureShare);
        GalleryDrawing.Fit(drawing, item.Plane, item.StackedFigureShare);
        KeepFitted(drawing);
        UpdateTour();
        Publish(item.Path, item.Title + " - " + AppTitle, push);
    }

    Drawing fittedDrawing;

    /// <summary>
    /// A drawing of the gallery that nobody touched yet keeps filling the window. Through the
    /// drawing's own event, not the canvas's: the coordinate system handles that one too (it
    /// keeps the middle in the middle), and it subscribed first, so this runs after it.
    /// </summary>
    void KeepFitted(Drawing drawing)
    {
        if (fittedDrawing != null)
        {
            fittedDrawing.SizeChanged -= FittedDrawing_SizeChanged;
        }

        fittedDrawing = drawing;
        if (fittedDrawing != null)
        {
            fittedDrawing.SizeChanged += FittedDrawing_SizeChanged;
        }
    }

    void FittedDrawing_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        var drawing = (Drawing)sender;
        if (CurrentSample != null && drawing == DrawingHost.CurrentDrawing && !drawing.ActionManager.CanUndo)
        {
            HandleExceptions(() => GalleryDrawing.Fit(drawing, CurrentSample.Plane, CurrentSample.StackedFigureShare));
        }
    }

    void ShowNeighborSample(int step)
    {
        if (CurrentSample == null)
        {
            return;
        }

        var items = GalleryCatalog.Items;
        int index = (GalleryCatalog.IndexOf(CurrentSample) + step + items.Count) % items.Count;
        ShowSample(items[index], push: true);
    }

    /// <summary>
    /// The drawing in the editor is the user's now (opened from a file, or a drawing of the
    /// gallery that they saved), from the named file
    /// </summary>
    /// <param name="file">The file for Save to write again, null to have it ask</param>
    void BecomeOwnDrawing(string fileName, IStorageFile file)
    {
        if (DrawingHost.CurrentDrawing != null)
        {
            // the user's own drawing: labels are theirs to move again, and "zoom to fit"
            // no longer lays out the caption (see OpenDrawing)
            DrawingHost.CurrentDrawing.FixedLabels.Clear();
            DrawingHost.CurrentDrawing.FitToWindow = null;
        }

        CurrentSample = null;
        OwnDrawing = null;
        OwnFileName = fileName;
        OwnFile = file;
        ReplaceKeptDrawing();
        UpdateTour();
        Publish(OwnDrawingPath, OwnTitle, push: true);
    }

    void UpdateTour()
    {
        // the toolbar is the editor's: made later, it asks itself (CreateEditor)
        if (!IsEditorBuilt)
        {
            return;
        }

        bool isSample = CurrentSample != null;
        TourGroup.IsVisible = isSample || OwnFileName != null;
        TourPrevious.IsVisible = isSample;
        TourPosition.IsVisible = isSample;
        TourNext.IsVisible = isSample;
        if (isSample)
        {
            TourPosition.Text = (GalleryCatalog.IndexOf(CurrentSample) + 1) + "/" + GalleryCatalog.Items.Count;
            TourTitle.Text = CurrentSample.Title;
        }
        else
        {
            TourTitle.Text = OwnFileName;
        }

        // the ribbon's default depends on the same thing
        UpdateRibbon();
    }

    #endregion

    async void OpenDrawingFromFile()
    {
        try
        {
            var topLevel = TopLevel.GetTopLevel(this);
            var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Open drawing",
                FileTypeFilter = new[] { AnyDrawingFileType, LgfFileType, DgfFileType, GgbFileType, FilePickerFileTypes.All }
            });

            var file = files?.FirstOrDefault();
            if (file == null)
            {
                return;
            }

            byte[] bytes;
            using (var stream = await file.OpenReadAsync())
            using (var memory = new MemoryStream())
            {
                await stream.CopyToAsync(memory);
                bytes = memory.ToArray();
            }

            OpenDrawing(file.Name, bytes, file);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    /// <summary>A drawing from the command line</summary>
    async void OpenDrawingFromPath(string path)
    {
        try
        {
            // the file as the storage provider knows it, for Save to write it again
            var file = await TopLevel.GetTopLevel(this).StorageProvider.TryGetFileFromPathAsync(Path.GetFullPath(path));
            OpenDrawing(Path.GetFileName(path), File.ReadAllBytes(path), file);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    /// <summary>
    /// A drawing to open at startup: the desktop head puts the file from its command line here
    /// </summary>
    public static string StartupFile { get; set; }

    /// <summary>A page to open at startup instead of the one the address bar says ("--gallery slug")</summary>
    public static string StartupPath { get; set; }

    /// <summary>"--arrange": the gallery's tiles can be dragged into a new order (see GalleryView)</summary>
    public static bool ArrangeGallery { get; set; }

    /// <summary>The first page: the file from the command line, else what the address says</summary>
    void OpenStartupFile()
    {
        AddressBar.Current.PathChanged += path => HandleExceptions(() => Navigate(path, push: false));

        if (CheckFolder != null)
        {
            RunCheck(CheckFolder, CheckOutputFolder);
            return;
        }

        if (ModernizeFolder != null)
        {
            RunModernize(ModernizeFolder);
            return;
        }

        if (RewriteFolder != null)
        {
            RunRewrite(RewriteFolder);
            return;
        }

        if (RecaptionFolder != null)
        {
            RunRecaption(RecaptionFolder);
            return;
        }

        if (SpaceLabelsFolder != null)
        {
            RunSpaceLabels(SpaceLabelsFolder);
            return;
        }

        var path = StartupFile;
        StartupFile = null;
        if (string.IsNullOrEmpty(path))
        {
            HandleExceptions(() => Navigate(StartupPath ?? AddressBar.Current.Path, push: false));
            return;
        }

        OpenDrawingFromPath(path);
    }

    /// <param name="name">File name; the extension tells the format</param>
    /// <param name="file">Where the bytes are from, if Save may write there again; only an .lgf is written again</param>
    public void OpenDrawing(string name, byte[] bytes, IStorageFile file = null)
    {
        // What the file is comes first: one that is no drawing leaves the page as it was.
        // (The page used to change hands before the file was read. A picture picked by
        // mistake left the drawing on screen under the picture's name - a drawing of the
        // gallery became the user's own that way, and one parked behind the gallery was
        // dropped.)
        bool isDgf = name.EndsWith(".dgf", StringComparison.OrdinalIgnoreCase);
        bool isGeoGebra = name.EndsWith(".ggb", StringComparison.OrdinalIgnoreCase);
        XElement xml = null;
        string text = null;
        string problem = null;
        if (isGeoGebra)
        {
            // a GeoGebra worksheet: a zip with the construction as geogebra.xml inside
            try
            {
                xml = GeoGebraReader.ReadWorksheet(bytes, out problem);
            }
            catch (Exception ex)
            {
                problem = "This file is damaged: " + ex.Message;
            }
        }
        else if (!isDgf)
        {
            text = Utilities.StripByteOrderMark(new System.Text.UTF8Encoding().GetString(bytes));
            xml = DrawingControl.ParseDrawing(text, out problem);
        }
        else if (!DecodeLegacyText(bytes).Contains("[General]", StringComparison.OrdinalIgnoreCase))
        {
            // every drawing DG wrote starts with its [General] section
            problem = "This file is not a drawing.";
        }

        if (problem != null)
        {
            RefuseFile(name, problem);
            return;
        }

        BecomeOwnDrawing(name, file: null);
        ShowEditor();

        if (isDgf)
        {
            // drawings of the original VB6 DG: INI-like text in the Windows ANSI code page
            var lines = DecodeLegacyText(bytes).Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            HandleExceptions(() => DrawingHost.DrawingControl.LoadDrawingFromDGF(lines, name));
            return;
        }

        if (isGeoGebra)
        {
            HandleExceptions(() => DrawingHost.DrawingControl.LoadDrawingFromGeoGebra(xml, name));
            return;
        }

        HandleExceptions(() =>
        {
            DrawingHost.DrawingControl.LoadDrawing(xml, name);
            var drawing = DrawingHost.CurrentDrawing;

            // a file that was not read in full has not given the drawing its name: Save
            // must not write what is left of it over the file
            if (drawing.Name == name)
            {
                OwnFile = file;
            }

            // a drawing with a caption (one of the gallery, saved) is laid out for this
            // window - once: the caption is the user's own to move now, and laid out again
            // by "zoom to fit" it would change what the file saves without an undo step
            if (GalleryDrawing.HasCaption(drawing))
            {
                GalleryDrawing.Fit(drawing, GalleryDrawing.GetPlane(text));
            }
        });
    }

    /// <summary>A file that can't be opened: said in the status bar, which only the editor has</summary>
    void RefuseFile(string name, string problem)
    {
        if (!LayoutRoot.IsVisible)
        {
            HandleExceptions(() => ShowOwnDrawing(push: true));
        }

        DrawingHost.ShowHint(name + ": " + problem);
    }

    /// <summary>
    /// VB6 wrote .dgf files in the system ANSI code page, which for DG's audience was
    /// mostly Cyrillic (1251). Valid UTF-8 is taken as is.
    /// </summary>
    static string DecodeLegacyText(byte[] bytes)
    {
        // checked, not tried: a DecoderFallbackException, even caught, is reported as an error
        // (CurrentDomain_FirstChanceException) for every file with Cyrillic in it
        if (System.Text.Unicode.Utf8.IsValid(bytes))
        {
            return new System.Text.UTF8Encoding(false).GetString(bytes);
        }

        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        return System.Text.Encoding.GetEncoding(1251).GetString(bytes);
    }

    /// <summary>
    /// Save: a drawing that has its .lgf file (<see cref="OwnFile"/>) is written there again,
    /// without a dialog. Anything else - a new drawing, one of the gallery, one read from a
    /// GeoGebra or DG file - has no file of this format yet and asks for one.
    /// </summary>
    async void SaveDrawing()
    {
        var file = OwnFile;
        if (CurrentSample != null || file == null)
        {
            SaveDrawingAs();
            return;
        }

        try
        {
            // a file that can't be written now (read-only, open in another program): it is
            // said why, and Save as offers another
            if (!await WriteDrawing(file))
            {
                SaveDrawingAs();
                return;
            }

            DrawingHost.ShowHint("Saved " + file.Name);
        }
        catch (Exception ex) when (IsReadOnlyFile(ex))
        {
            SaveDrawingAs();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    /// <summary>
    /// A browser without the File System Access API (Firefox, Safari) opens a file to read
    /// and has nothing to write it with: there Save is a download under a name, every time
    /// </summary>
    static bool IsReadOnlyFile(Exception exception)
    {
        return exception.GetType().Name == "JSException" && exception.Message.Contains("not a writeable file");
    }

    /// <summary>Save as: asks where to, and that file is the drawing's from then on</summary>
    async void SaveDrawingAs()
    {
        try
        {
            var name = CurrentSample != null ? CurrentSample.FileName : OwnFileName;
            var topLevel = TopLevel.GetTopLevel(this);
            var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Save drawing",
                SuggestedFileName = Path.ChangeExtension(name ?? "drawing", "lgf"),
                DefaultExtension = "lgf",
                FileTypeChoices = new[] { LgfFileType }
            });

            if (file == null)
            {
                return;
            }

            if (!await WriteDrawing(file))
            {
                return;
            }

            // a saved drawing of the gallery is the user's own from here on; a drawing of
            // their own goes by its new name
            if (CurrentSample != null)
            {
                BecomeOwnDrawing(file.Name, file);
            }
            else
            {
                OwnFileName = file.Name;
                OwnFile = file;
                KeepOwnDrawingSoon();
                UpdateTour();
                AddressBar.Current.Replace(OwnDrawingPath, OwnTitle);
            }

            DrawingHost.ShowHint("Saved " + file.Name);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    /// <returns>Whether the file was written (see <see cref="TryWriteFile"/>)</returns>
    async Task<bool> WriteDrawing(IStorageFile file)
    {
        // A construction under way is not in the drawing yet: its point following the
        // cursor and its preview are figures like any other while it lasts, and went into
        // the file ("TempPoint", half a polygon). It is put away, as Escape does.
        var drawing = DrawingHost.CurrentDrawing;
        if (DrawingHost.DrawingControl.ConstructionInProgress || drawing.IsRecordingTransaction)
        {
            drawing.Behavior?.Restart();
        }

        return await TryWriteFile(file, new System.Text.UTF8Encoding(false).GetBytes(drawing.SaveAsText()));
    }

    // a file is being written: what the system says against it is said in words, not reported
    static bool writingFile;

    /// <summary>
    /// The bytes into the file. One that can't be written - read-only, open in another
    /// program, in a folder one may not write to - is said so in the hint, and false comes
    /// back. (It was an error report, as for a bug, and the system's message in the hint.)
    /// </summary>
    async Task<bool> TryWriteFile(IStorageFile file, byte[] bytes)
    {
        writingFile = true;
        try
        {
            // bytes straight into the stream: the browser's file stream only has WriteAsync,
            // and a StreamWriter flushes synchronously when disposed
            await using (var stream = await file.OpenWriteAsync())
            {
                await stream.WriteAsync(bytes);
                await stream.FlushAsync();
            }

            return true;
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
        {
            DrawingHost.ShowHint("Could not write " + file.Name + ": it may be read-only or open in another program.");
            return false;
        }
        finally
        {
            writingFile = false;
        }
    }

    void Copy() => HandleExceptions(() => DrawingHost.CurrentDrawing.Copy());

    void Paste() => HandleExceptions(() => DrawingHost.CurrentDrawing.Paste());

    private void DeleteSelection() => HandleExceptions(() => DrawingHost.CurrentDrawing.DeleteSelection());

    private void SelectAll() => HandleExceptions(() =>
    {
        var drawing = DrawingHost.CurrentDrawing;
        drawing.SelectAll();
        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
    });

    // Had a menu item until the menu went away; no way to reach them for now:
    // "Lock" (Drawing.LockSelected) and the settings page below.

    #region Settings

    Settings PageSettings;

    class Settings
    {
        MainView Page;

        public Settings(MainView page)
        {
            Page = page;
        }

        [PropertyGridVisible]
        [PropertyGridName("Show coordinate axes and grid")]
        public bool ShowGrid
        {
            get => Page.DrawingHost.CurrentDrawing.CoordinateGrid.Visible;
            set => Page.DrawingHost.CurrentDrawing.CoordinateGrid.Visible = value;
        }
    }

    private void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        PageSettings ??= new Settings(this);
        if (DrawingHost.PropertyGrid.Selection == PageSettings)
        {
            DrawingHost.ShowProperties(null);
        }
        else
        {
            DrawingHost.ShowProperties(PageSettings);
        }
    }

    #endregion

    const double KeyboardPanPixels = 48;

    /// <summary>
    /// Keys without modifiers: tool letters, arrows to pan, +/- to zoom, H to see everything.
    /// </summary>
    bool HandlePlainKey(Key key)
    {
        var drawing = DrawingHost.CurrentDrawing;
        if (drawing == null)
        {
            return false;
        }

        var toolType = BehaviorShortcuts.GetTool(key);
        if (toolType != null)
        {
            var tool = Behaviors.FirstOrDefault(b => b.GetType() == toolType);
            if (tool == null)
            {
                return false;
            }

            drawing.Behavior = tool;
            return true;
        }

        var coordinateSystem = drawing.CoordinateSystem;
        switch (key)
        {
            case Key.Left: Pan(KeyboardPanPixels, 0); return true;
            case Key.Right: Pan(-KeyboardPanPixels, 0); return true;
            case Key.Up: Pan(0, KeyboardPanPixels); return true;
            case Key.Down: Pan(0, -KeyboardPanPixels); return true;
            case Key.Add:
            case Key.OemPlus:
                coordinateSystem.ZoomIn();
                return true;
            case Key.Subtract:
            case Key.OemMinus:
                coordinateSystem.ZoomOut();
                return true;
            case Key.G:
                DrawingHost.ToggleGrid();
                return true;
            case Key.H:
                HandleExceptions(() => drawing.ZoomToFit());
                return true;
            case Key.Home:
                // (a drawing laid out around its caption has its middle elsewhere than its content)
                HandleExceptions(() =>
                {
                    if (drawing.FitToWindow != null)
                    {
                        drawing.ZoomToFit();
                    }
                    else
                    {
                        coordinateSystem.CenterContent();
                    }
                });
                return true;
            case Key.PageUp:
                HandleExceptions(() => ShowNeighborSample(-1));
                return CurrentSample != null;
            case Key.PageDown:
                HandleExceptions(() => ShowNeighborSample(1));
                return CurrentSample != null;
        }

        return false;

        void Pan(double physicalX, double physicalY)
        {
            var offset = new Avalonia.Point(
                coordinateSystem.ToLogical(physicalX),
                -coordinateSystem.ToLogical(physicalY));

            // the same list for every press, so that a run of arrow keys is one undo step,
            // as a drag of the canvas is (moves of the same list merge)
            if (keyboardPan == null || keyboardPan[0] != coordinateSystem)
            {
                keyboardPan = new IMovable[] { coordinateSystem };
            }

            // in the middle of a construction the view just moves: recorded, the pan would
            // be part of the figure's undo step, and Escape would take the view back
            if (drawing.IsRecordingTransaction)
            {
                keyboardPan.Move(offset);
                return;
            }

            // the same undoable move that dragging the empty canvas does
            Actions.Move(drawing, keyboardPan, offset, toRecalculate: null);
        }
    }

    IMovable[] keyboardPan;

    private void MainView_KeyUp(object sender, KeyEventArgs e)
    {
        if (IsGalleryShowing || !IsEditorBuilt)
        {
            shortcutKeyDown = Key.None;
            return;
        }

        if (!IsPlainKeyForCanvas(e.Key))
        {
            return;
        }

        // handled on the way down (MainView_KeyDown), where a held key repeats
        if (IsRepeatingKey(e.Key))
        {
            return;
        }

        // the letter of a Ctrl shortcut, let go after Ctrl: not a tool letter
        if (e.Key == shortcutKeyDown)
        {
            shortcutKeyDown = Key.None;
            return;
        }

        if (e.KeyModifiers == KeyModifiers.None && HandlePlainKey(e.Key))
        {
            e.Handled = true;
            return;
        }

        // a Mac's "delete" key is Backspace
        if (e.Key == Key.Delete || e.Key == Key.Back)
        {
            DeleteSelection();
        }
    }

    Key shortcutKeyDown = Key.None;

    /// <summary>
    /// Whether a key without modifiers is the canvas's (<see cref="HandlePlainKey"/>) rather
    /// than that of the control with the keyboard. Both key handlers tunnel, so they see the
    /// key before that control does.
    /// </summary>
    bool IsPlainKeyForCanvas(Key key)
    {
        var focused = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement();
        if (focused is TextBox)
        {
            return false;
        }

        // up and down the Figure List, not panning the canvas
        if (DrawingHost.FigureExplorer.IsKeyboardFocusWithin && FigureExplorer.IsNavigationKey(key))
        {
            return false;
        }

        // Nor from a list, a slider or a combo of the side panel, which takes the key on its
        // way down: an arrow that picked the next style also moved the view by a step, as an
        // undo step. (Tool letters still work from there.)
        if (DrawingHost.PropertyGrid.IsKeyboardFocusWithin && IsViewKey(key))
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// The plain keys that act on key down, so that holding one repeats it: the arrows pan
    /// and +/- zoom on. The others act on key up, once (a held G would flicker the grid, a
    /// held Page Down would race through the gallery).
    /// </summary>
    static bool IsRepeatingKey(Key key)
    {
        switch (key)
        {
            case Key.Left:
            case Key.Right:
            case Key.Up:
            case Key.Down:
            case Key.Add:
            case Key.OemPlus:
            case Key.Subtract:
            case Key.OemMinus:
                return true;
        }

        return false;
    }

    /// <summary>The keys that move the view (<see cref="HandlePlainKey"/>), which controls with a selection or a value use too</summary>
    static bool IsViewKey(Key key)
    {
        switch (key)
        {
            case Key.Left:
            case Key.Right:
            case Key.Up:
            case Key.Down:
            case Key.Home:
            case Key.PageUp:
            case Key.PageDown:
            case Key.Add:
            case Key.OemPlus:
            case Key.Subtract:
            case Key.OemMinus:
                return true;
        }

        return false;
    }

    #region Empty chrome

    /// <summary>
    /// A press on the chrome where there is nothing to press - the empty stretch of the ribbon,
    /// the toolbar right of its buttons, the Figure List below its rows - puts the side panel
    /// away (<see cref="DrawingHost.CloseSidePanel"/>). On a phone the panel can cover most of
    /// the screen, with no empty bit of canvas left to click.
    /// </summary>
    void MainView_PointerPressed(object sender, PointerPressedEventArgs e)
    {
        if (IsEditorBuilt && DrawingHost.IsSidePanelShown && e.Source is Avalonia.Visual source && IsEmptyChrome(source))
        {
            HandleExceptions(DrawingHost.CloseSidePanel);
        }
    }

    /// <summary>
    /// From what was hit up to the ribbon, the toolbar or the Figure List through nothing but
    /// layout: a button, a tab, a row, the splitter or a scroll bar on the way is something to press
    /// </summary>
    bool IsEmptyChrome(Avalonia.Visual visual)
    {
        for (var current = visual; current != null; current = Avalonia.VisualTree.VisualExtensions.GetVisualParent(current))
        {
            if (current == DrawingHost.Ribbon || current == Toolbar || current == DrawingHost.FigureExplorer)
            {
                return true;
            }

            if (!IsLayoutOnly(current))
            {
                return false;
            }
        }

        return false;
    }

    /// <summary>Parts that only lay out or draw; subclasses (a toolbar button is a Border) are more than that</summary>
    static bool IsLayoutOnly(Avalonia.Visual visual)
    {
        var type = visual.GetType();
        return type == typeof(Border)
            || type == typeof(Panel)
            || type == typeof(StackPanel)
            || type == typeof(DockPanel)
            || type == typeof(WrapPanel)
            || type == typeof(Grid)
            || type == typeof(TextBlock)
            || type == typeof(Avalonia.Controls.Presenters.ContentPresenter)
            || type == typeof(Avalonia.Controls.Presenters.ItemsPresenter)
            || type == typeof(Avalonia.Controls.Presenters.ScrollContentPresenter)
            || type == typeof(ScrollViewer)
            || type == typeof(MainToolbarGroup)
            || visual is Avalonia.Controls.Shapes.Shape;
    }

    #endregion

    /// <summary>
    /// The modifier of the shortcuts: Ctrl, or Cmd on a Mac (the browser app there, where
    /// Ctrl+click is a right click and Cmd+S, ours ignored, saved the web page)
    /// </summary>
    static bool IsCommandModifier(KeyModifiers modifiers)
    {
        return modifiers == KeyModifiers.Control || modifiers == KeyModifiers.Meta;
    }

    /// <summary>Ctrl+Shift+Z (Cmd+Shift+Z on a Mac): redo, as most programs have it besides Ctrl+Y</summary>
    static bool IsRedoChord(KeyModifiers modifiers, Key key)
    {
        return key == Key.Z && IsCommandModifier(modifiers & ~KeyModifiers.Shift) && modifiers.HasFlag(KeyModifiers.Shift);
    }

    /// <summary>
    /// Ctrl+letter, on the way down. (On the way up the state of Ctrl depends on which of the
    /// two keys was let go first, and a bare "S" is the Segment tool.)
    /// </summary>
    bool HandleControlShortcut(Key key)
    {
        switch (key)
        {
            case Key.Z: DrawingHost.DrawingControl.Undo(); return true;
            case Key.Y: DrawingHost.DrawingControl.Redo(); return true;
            case Key.A: SelectAll(); return true;
            case Key.N: NewDrawing(); return true;
            case Key.O: OpenDrawingFromFile(); return true;
            case Key.S: SaveDrawing(); return true;
            case Key.C: Copy(); return true;
            case Key.V: Paste(); return true;
            case Key.F1: ToggleRibbon(); return true; // as in Office
        }

        return false;
    }

    /// <summary>
    /// Escape aborts the construction in progress; with nothing in progress it goes back to
    /// the default tool. Handled here, on the way down and regardless of focus: tools with an
    /// input panel (Point by coordinates...) keep pulling the focus into their text box, where
    /// neither the canvas nor the KeyUp handler below would ever see the key.
    /// </summary>
    private void MainView_KeyDown(object sender, KeyEventArgs e)
    {
        if (IsGalleryShowing)
        {
            // there is no drawing to undo, select or save
            if (IsCommandModifier(e.KeyModifiers) && (e.Key == Key.N || e.Key == Key.O))
            {
                HandleControlShortcut(e.Key);
                shortcutKeyDown = e.Key;
                e.Handled = true;
            }

            return;
        }

        if (!IsEditorBuilt || DrawingHost.CurrentDrawing == null)
        {
            return;
        }

        // A plain key pressed anew is no longer the letter of a Ctrl shortcut. (The release
        // of that letter is skipped on its way up - but after Ctrl+S or Ctrl+O it goes to
        // the file dialog and never arrives, and the next S, for the Segment tool, was the
        // one skipped.)
        if (e.KeyModifiers == KeyModifiers.None && e.Key == shortcutKeyDown)
        {
            shortcutKeyDown = Key.None;
        }

        bool redo = IsRedoChord(e.KeyModifiers, e.Key);
        if (IsCommandModifier(e.KeyModifiers) || redo)
        {
            // In a text box Ctrl+C, V, X, Z, Y, A are the text box's own. Save, Open, New
            // and the ribbon are not: with the keyboard in a box (where a tool's panel keeps
            // putting it) Ctrl+S did nothing, and in the browser the page got it. What is
            // typed in the box goes into the drawing first, as when the box is left.
            var focused = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement();
            if (focused is TextBox)
            {
                if (e.Key != Key.S && e.Key != Key.O && e.Key != Key.N && e.Key != Key.F1)
                {
                    return;
                }

                DrawingHost.DrawingControl.Focus();
            }

            if (redo)
            {
                DrawingHost.DrawingControl.Redo();
                shortcutKeyDown = e.Key;
                e.Handled = true;
            }
            else if (HandleControlShortcut(e.Key))
            {
                shortcutKeyDown = e.Key;
                e.Handled = true;
            }

            return;
        }

        // on the way down, where a held key repeats (on the way up it panned once)
        if (e.KeyModifiers == KeyModifiers.None && IsRepeatingKey(e.Key))
        {
            if (IsPlainKeyForCanvas(e.Key) && HandlePlainKey(e.Key))
            {
                e.Handled = true;
            }

            return;
        }

        if (e.Key != Key.Escape)
        {
            return;
        }

        var behavior = DrawingHost.CurrentDrawing.Behavior;
        if (behavior.IsInInitialState)
        {
            DrawingHost.CurrentDrawing.SetDefaultBehavior();
        }
        else
        {
            behavior.Restart();

            // back to how the tool looks when freshly picked
            DrawingHost.ShowHint(behavior.HintText);
            DrawingHost.ShowProperties(behavior.PropertyBag);
        }

        DrawingHost.DrawingControl.Focus();
        e.Handled = true;
    }

    /// <summary>
    /// A key pressed while nothing has the focus. That is the state after the focused control
    /// is taken out of the tree (the Fix length button, when the length panel is rebuilt
    /// around it): Avalonia gives the focus to no one, and a key then goes to the window
    /// alone, past MainView and the canvas - Escape, tool letters, Ctrl shortcuts all dead.
    /// Gives the focus back to the canvas (to this view on the gallery) and sends the key
    /// on from there.
    /// </summary>
    private void TopLevel_KeyDown(object sender, KeyEventArgs e)
    {
        var topLevel = (TopLevel)sender;
        if (topLevel.FocusManager?.GetFocusedElement() != null)
        {
            return;
        }

        Control target = drawingHost?.DrawingControl;
        if (target == null || !target.Focus())
        {
            target = this;
            if (!target.Focus())
            {
                return;
            }
        }

        var forwarded = new KeyEventArgs()
        {
            RoutedEvent = e.RoutedEvent,
            Key = e.Key,
            KeyModifiers = e.KeyModifiers,
            PhysicalKey = e.PhysicalKey,
            KeySymbol = e.KeySymbol,
            KeyDeviceType = e.KeyDeviceType
        };
        target.RaiseEvent(forwarded);
        e.Handled = forwarded.Handled;
    }
}
