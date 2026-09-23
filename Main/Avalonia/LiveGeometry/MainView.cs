using System;
using System.IO;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;
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
    DrawingHost DrawingHost = new DrawingHost();

    Behavior[] Behaviors = Array.Empty<Behavior>();

    static readonly FilePickerFileType LgfFileType = new("Live Geometry drawing")
    {
        Patterns = new[] { "*.lgf" }
    };

    static readonly FilePickerFileType AnyDrawingFileType = new("All drawings")
    {
        Patterns = new[] { "*.lgf", "*.dgf" }
    };

    static readonly FilePickerFileType DgfFileType = new("DG 1.x drawing")
    {
        Patterns = new[] { "*.dgf" }
    };

    public MainView()
    {
        InitializeComponent();
        AddBehaviors();
        LayoutRoot.Children.Add(DrawingHost);
        InitializeCommands();

        // The geometry library surfaces errors through the WPF-style MessageBox shim.
        MessageBox.Handler = text => DrawingHost.ShowHint(text);
        DrawingHost.UnhandledException += (s, e) =>
        {
            Console.WriteLine("LiveGeometry error: " + e.Exception);
            DrawingHost.ShowHint(e.Exception.Message);
        };

        // Give the drawing canvas keyboard focus so behaviors receive Escape/Delete/etc.
        DrawingHost.DrawingControl.Focusable = true;
        DrawingHost.DrawingControl.PointerPressed += (s, e) => DrawingHost.DrawingControl.Focus();

        AddHandler(KeyDownEvent, MainView_KeyDown, RoutingStrategies.Tunnel);
        AddHandler(KeyUpEvent, MainView_KeyUp, RoutingStrategies.Tunnel);

        // a phone turned on its side is a different screen
        SizeChanged += (s, e) => UpdateRibbon();

        AttachedToVisualTree += (s, e) =>
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel != null)
            {
                DynamicGeometry.Clipboard.SystemClipboard = topLevel.Clipboard;
            }

            Focus();

            // once the canvas has a size, or the drawing would be laid out in a 0x0 viewport
            Avalonia.Threading.Dispatcher.UIThread.Post(OpenStartupFile, Avalonia.Threading.DispatcherPriority.Loaded);
        };
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
    }

    // A link to the repository at the far right of the toolbar; its tooltip is the build (git
    // commit) on screen - e.g. whether a fresh deployment has arrived yet. A faint Octocat
    // rather than the raw commit hash, which meant nothing to the kids the app is for.
    const string RepositoryUrl = "https://github.com/KirillOsenkov/LiveGeometry";

    const double OctocatSize = 16;
    const double OctocatOpacity = 0.4;

    // the GitHub mark (the "mark-github" octicon, on a 16 x 16 grid)
    const string OctocatPath = "M8,0 C3.58,0 0,3.58 0,8 c0,3.54 2.29,6.53 5.47,7.59 c0.4,0.07 0.55,-0.17 0.55,-0.38 c0,-0.19 -0.01,-0.82 -0.01,-1.49 c-2.01,0.37 -2.53,-0.49 -2.69,-0.94 c-0.09,-0.23 -0.48,-0.94 -0.82,-1.13 c-0.28,-0.15 -0.68,-0.52 -0.01,-0.53 c0.63,-0.01 1.08,0.58 1.23,0.82 c0.72,1.21 1.87,0.87 2.33,0.66 c0.07,-0.52 0.28,-0.87 0.51,-1.07 c-1.78,-0.2 -3.64,-0.89 -3.64,-3.95 c0,-0.87 0.31,-1.59 0.82,-2.15 c-0.08,-0.2 -0.36,-1.02 0.08,-2.12 c0,0 0.67,-0.21 2.2,0.82 c0.64,-0.18 1.32,-0.27 2,-0.27 c0.68,0 1.36,0.09 2,0.27 c1.53,-1.04 2.2,-0.82 2.2,-0.82 c0.44,1.1 0.16,1.92 0.08,2.12 c0.51,0.56 0.82,1.27 0.82,2.15 c0,3.07 -1.87,3.75 -3.65,3.95 c0.29,0.25 0.54,0.73 0.54,1.48 c0,1.07 -0.01,1.93 -0.01,2.2 c0,0.21 0.15,0.46 0.55,0.38 A8.013,8.013 0 0 0 16,8 c0,-4.42 -3.58,-8 -8,-8 z";

    static Control CreateBuildStamp()
    {
        var octocat = new Border()
        {
            Background = Brushes.Transparent, // hit-testable around the cat too
            Cursor = new Cursor(StandardCursorType.Hand),
            Opacity = OctocatOpacity,
            Margin = new Avalonia.Thickness(8, 0, 10, 0),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            Child = new Avalonia.Controls.Shapes.Path()
            {
                Data = Geometry.Parse(OctocatPath),
                Fill = RibbonTheme.Text,
                Stretch = Stretch.Uniform,
                Width = OctocatSize,
                Height = OctocatSize
            }
        };
        ToolTip.SetTip(octocat, BuildVersion.Full + "\n" + RepositoryUrl);
        octocat.PointerEntered += (s, e) => octocat.Opacity = 1;
        octocat.PointerExited += (s, e) => octocat.Opacity = OctocatOpacity;
        octocat.PointerReleased += (s, e) =>
        {
            if (e.InitialPressMouseButton == MouseButton.Left)
            {
                TopLevel.GetTopLevel(octocat)?.Launcher.LaunchUriAsync(new Uri(RepositoryUrl));
            }
        };
        return octocat;
    }

    private void InitializeComponent()
    {
        Focusable = true;
        Console.WriteLine("Live Geometry " + BuildVersion.Full);

        // Two pages, one showing: the gallery (the start page) and the editor. Started with a
        // file (or a batch job) the editor is up from the first frame and the gallery, with its
        // 47 tiles, isn't even built until the Gallery button is pressed.
        pages.Children.Add(LayoutRoot);
        Content = pages;
        bool startsInEditor = StartupFile != null || CheckFolder != null || ModernizeFolder != null;
        LayoutRoot.IsVisible = startsInEditor;
        if (!startsInEditor)
        {
            EnsureGallery();
        }

        // No menu: the few document commands are a toolbar, everything else is the keyboard
        // (see MainView_KeyUp and HandlePlainKey), the mouse wheel and the context menu.
        var toolbar = Toolbar;
        toolbar.AddAtRight(CreateBuildStamp());
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
        toolbar.AddButton(MainToolbarIcons.Save(), "Save", "Ctrl+S", SaveDrawingToFile);
        toolbar.AddSeparator();
        toolbar.AddButton(MainToolbarIcons.Undo(), "Ctrl+Z", DrawingHost.DrawingControl.CommandUndo);
        toolbar.AddButton(MainToolbarIcons.Redo(), "Ctrl+Y", DrawingHost.DrawingControl.CommandRedo);

        // only while a drawing of the gallery is open: previous / next through the gallery,
        // in the middle of the room the toolbar has left, and bigger than the document buttons
        TourGroup = toolbar.BeginCenteredGroup();
        toolbar.AddButton(
            MainToolbarIcons.Previous(),
            "Previous drawing",
            "Page Up",
            () => HandleExceptions(() => ShowNeighborSample(-1)),
            iconSize: TourIconSize,
            inset: TourArrowInset);
        TourPosition = toolbar.AddText(FontWeight.Normal, minWidth: 44, fontSize: TourFontSize);
        toolbar.AddButton(
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
        UpdateRibbon();
    }

    #region Ribbon

    // The ribbon can be folded away behind the first button of the toolbar, to leave a small
    // screen to the drawing. It starts folded when a drawing of the gallery is opened on a small
    // screen (one is there to look, and the tabs wouldn't fit anyway) and open everywhere else;
    // once the user has pressed the button, their choice holds for the rest of the session.

    readonly MainToolbar Toolbar = new MainToolbar();
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
        bool visible = ribbonChoice ?? !(CurrentSample != null && IsSmallScreen);
        DrawingHost.Ribbon.IsVisible = visible;
        Toolbar.IsTabOpen = visible;
        ToolTip.SetTip(ToolboxButton, (visible ? "Hide the tools" : "Show the tools") + " (" + RibbonShortcut + ")");
    }

    #endregion

    void InitializeCommands()
    {
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleGrid, first: true);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleOrtho);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleSnapToGrid);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleSnapToPoint);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleSnapToCenter);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleLabelNewPoints);
        DrawingHost.AddToolbarButton(DrawingHost.CommandTogglePolar);
        DrawingHost.AddToolbarButton(DrawingHost.CommandTogglePointByCoordinates);
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

        Gallery = new GalleryView(CreateBuildStamp());
        Gallery.NewDrawingRequested += () => HandleExceptions(() => ShowNewDrawing(push: true));
        Gallery.ContinueDrawingRequested += () => HandleExceptions(() => ShowOwnDrawing(push: true));
        Gallery.ItemRequested += item => HandleExceptions(() => ShowSample(item, push: true));
        pages.Children.Add(Gallery);
    }

    const double TourIconSize = 32;
    const double TourFontSize = 17;
    const double TourArrowInset = 1; // the chevrons have room enough inside their own icon

    Panel TourGroup;
    TextBlock TourPosition;
    TextBlock TourTitle;

    /// <summary>The drawing of the gallery that is open, null for a drawing of the user's own</summary>
    GalleryItem CurrentSample;

    /// <summary>
    /// The user's own drawing, kept (with its undo history) while they look around the gallery
    /// </summary>
    Drawing OwnDrawing;

    void Navigate(string path, bool push)
    {
        var item = GalleryCatalog.FindByPath(path);
        if (item != null)
        {
            ShowSample(item, push);
        }
        else if (string.Equals(path.TrimEnd('/'), OwnDrawingPath, StringComparison.OrdinalIgnoreCase))
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
        var drawing = DrawingHost.CurrentDrawing;
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
        Gallery.CanContinueDrawing = OwnDrawing != null;
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
        DrawingHost.Clear();
        UpdateTour();
        Publish(OwnDrawingPath, AppTitle, push);
    }

    /// <summary>Back to the drawing the user left for the gallery; a new one if there is none</summary>
    void ShowOwnDrawing(bool push)
    {
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
            // the same order as DrawingControl.Clear: on the canvas first, then the current one
            control.Background = Avalonia.Media.Brushes.White;
            OwnDrawing.Canvas = control;
            control.Drawing = OwnDrawing;
            OwnDrawing.Recalculate();
        }

        UpdateTour();
        Publish(OwnDrawingPath, AppTitle, push);
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
        GalleryDrawing.Fit(control.Drawing, item.Plane);
        KeepFitted(control.Drawing);
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
            HandleExceptions(() => GalleryDrawing.Fit(drawing, CurrentSample.Plane));
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
    /// gallery that they saved)
    /// </summary>
    void BecomeOwnDrawing()
    {
        CurrentSample = null;
        OwnDrawing = null;
        UpdateTour();
        Publish(OwnDrawingPath, AppTitle, push: true);
    }

    void UpdateTour()
    {
        TourGroup.IsVisible = CurrentSample != null;
        if (CurrentSample != null)
        {
            TourPosition.Text = (GalleryCatalog.IndexOf(CurrentSample) + 1) + "/" + GalleryCatalog.Items.Count;
            TourTitle.Text = CurrentSample.Title;
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
                FileTypeFilter = new[] { AnyDrawingFileType, LgfFileType, DgfFileType, FilePickerFileTypes.All }
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

            OpenDrawing(file.Name, bytes);
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

        var path = StartupFile;
        StartupFile = null;
        if (string.IsNullOrEmpty(path))
        {
            HandleExceptions(() => Navigate(AddressBar.Current.Path, push: false));
            return;
        }

        HandleExceptions(() => OpenDrawing(Path.GetFileName(path), File.ReadAllBytes(path)));
    }

    /// <param name="name">File name; the extension tells the format</param>
    public void OpenDrawing(string name, byte[] bytes)
    {
        BecomeOwnDrawing();
        ShowEditor();

        if (name.EndsWith(".dgf", StringComparison.OrdinalIgnoreCase))
        {
            // drawings of the original VB6 DG: INI-like text in the Windows ANSI code page
            var lines = DecodeLegacyText(bytes).Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            HandleExceptions(() => DrawingHost.DrawingControl.LoadDrawingFromDGF(lines, name));
            return;
        }

        var text = Utilities.StripByteOrderMark(new System.Text.UTF8Encoding().GetString(bytes));
        HandleExceptions(() => DrawingHost.DrawingControl.LoadDrawing(text, name));
    }

    /// <summary>
    /// VB6 wrote .dgf files in the system ANSI code page, which for DG's audience was
    /// mostly Cyrillic (1251). Valid UTF-8 is taken as is.
    /// </summary>
    static string DecodeLegacyText(byte[] bytes)
    {
        try
        {
            return new System.Text.UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes);
        }
        catch (System.Text.DecoderFallbackException)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            return System.Text.Encoding.GetEncoding(1251).GetString(bytes);
        }
    }

    async void SaveDrawingToFile()
    {
        try
        {
            var topLevel = TopLevel.GetTopLevel(this);
            var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Save drawing",
                SuggestedFileName = CurrentSample != null ? CurrentSample.FileName : "drawing.lgf",
                DefaultExtension = "lgf",
                FileTypeChoices = new[] { LgfFileType }
            });

            if (file == null)
            {
                return;
            }

            var text = DrawingHost.CurrentDrawing.SaveAsText();
            using (var stream = await file.OpenWriteAsync())
            using (var writer = new StreamWriter(stream))
            {
                await writer.WriteAsync(text);
            }

            // a saved drawing of the gallery is the user's own from here on
            if (CurrentSample != null)
            {
                BecomeOwnDrawing();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    void Copy() => HandleExceptions(() => DrawingHost.CurrentDrawing.Copy());

    void Paste() => HandleExceptions(() => DrawingHost.CurrentDrawing.Paste());

    private void DeleteSelection() => HandleExceptions(() => DrawingHost.CurrentDrawing.DeleteSelection());

    private void SelectAll() => HandleExceptions(() => DrawingHost.CurrentDrawing.SelectAll());

    // Had a menu item until the menu went away; no way to reach them for now:
    // "Lock" (Drawing.LockSelected), "Figure List" (DrawingHost.CommandShowFigureExplorer)
    // and the settings page below.

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
            case Key.H:
                HandleExceptions(() => coordinateSystem.ZoomExtend());
                return true;
            case Key.Home:
                HandleExceptions(() => coordinateSystem.CenterContent());
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
            // the same undoable move that dragging the empty canvas does
            var offset = new Avalonia.Point(
                coordinateSystem.ToLogical(physicalX),
                -coordinateSystem.ToLogical(physicalY));
            Actions.Move(drawing, new IMovable[] { coordinateSystem }, offset, null);
        }
    }

    private void MainView_KeyUp(object sender, KeyEventArgs e)
    {
        if (IsGalleryShowing)
        {
            shortcutKeyDown = Key.None;
            return;
        }

        var focused = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement();
        if (focused is TextBox)
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

        if (e.Key == Key.Delete)
        {
            DeleteSelection();
        }
    }

    Key shortcutKeyDown = Key.None;

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
            case Key.S: SaveDrawingToFile(); return true;
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
            if (e.KeyModifiers == KeyModifiers.Control && (e.Key == Key.N || e.Key == Key.O))
            {
                HandleControlShortcut(e.Key);
                shortcutKeyDown = e.Key;
                e.Handled = true;
            }

            return;
        }

        if (DrawingHost.CurrentDrawing == null)
        {
            return;
        }

        if (e.KeyModifiers == KeyModifiers.Control)
        {
            // in a text box Ctrl+C, V, Z, A are the text box's own
            var focused = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement();
            if (!(focused is TextBox) && HandleControlShortcut(e.Key))
            {
                shortcutKeyDown = e.Key;
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
}
