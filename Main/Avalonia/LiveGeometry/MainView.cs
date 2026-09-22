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

        // a drawing of the gallery that nobody touched yet keeps filling the window
        DrawingHost.DrawingControl.SizeChanged += (s, e) =>
        {
            var drawing = DrawingHost.CurrentDrawing;
            if (CurrentSample != null && drawing != null && !drawing.ActionManager.CanUndo)
            {
                HandleExceptions(() => GalleryDrawing.Fit(drawing, CurrentSample.Plane));
            }
        };

        AddHandler(KeyDownEvent, MainView_KeyDown, RoutingStrategies.Tunnel);
        AddHandler(KeyUpEvent, MainView_KeyUp, RoutingStrategies.Tunnel);

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

    // The build (git commit) at the far right of the toolbar, so that it is obvious
    // which version is on screen - e.g. whether a fresh deployment has arrived yet.
    static Control CreateBuildStamp()
    {
        var build = new TextBlock()
        {
            Text = BuildVersion.Short,
            FontSize = 11,
            Opacity = 0.55,
            Margin = new Avalonia.Thickness(8, 0, 10, 0),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
        };
        ToolTip.SetTip(build, BuildVersion.Full);
        return build;
    }

    private void InitializeComponent()
    {
        Focusable = true;
        Console.WriteLine("Live Geometry " + BuildVersion.Full);

        // Two pages, one showing: the gallery (the start page) and the editor.
        Gallery = new GalleryView(CreateBuildStamp());
        Gallery.NewDrawingRequested += () => HandleExceptions(() => ShowNewDrawing(push: true));
        Gallery.ContinueDrawingRequested += () => HandleExceptions(() => ShowOwnDrawing(push: true));
        Gallery.ItemRequested += item => HandleExceptions(() => ShowSample(item, push: true));
        LayoutRoot.IsVisible = false;

        var pages = new Panel();
        pages.Children.Add(LayoutRoot);
        pages.Children.Add(Gallery);
        Content = pages;

        // No menu: the few document commands are a toolbar, everything else is the keyboard
        // (see MainView_KeyUp and HandlePlainKey), the mouse wheel and the context menu.
        var toolbar = new MainToolbar();
        toolbar.AddAtRight(CreateBuildStamp());
        toolbar.AddButton(MainToolbarIcons.Gallery(), "Gallery", shortcut: null, () => HandleExceptions(() => ShowGallery(push: true)));
        toolbar.AddSeparator();
        toolbar.AddButton(MainToolbarIcons.New(), "New", "Ctrl+N", NewDrawing);
        toolbar.AddButton(MainToolbarIcons.Open(), "Open", "Ctrl+O", OpenDrawingFromFile);
        toolbar.AddButton(MainToolbarIcons.Save(), "Save", "Ctrl+S", SaveDrawingToFile);
        toolbar.AddSeparator();
        toolbar.AddButton(MainToolbarIcons.Undo(), "Ctrl+Z", DrawingHost.DrawingControl.CommandUndo);
        toolbar.AddButton(MainToolbarIcons.Redo(), "Ctrl+Y", DrawingHost.DrawingControl.CommandRedo);

        // only while a drawing of the gallery is open: previous / next through the gallery
        TourGroup = toolbar.BeginGroup();
        toolbar.AddSeparator();
        toolbar.AddButton(MainToolbarIcons.Previous(), "Previous drawing", "Page Up", () => HandleExceptions(() => ShowNeighborSample(-1)));
        TourPosition = toolbar.AddText(FontWeight.Normal, minWidth: 52);
        toolbar.AddButton(MainToolbarIcons.Next(), "Next drawing", "Page Down", () => HandleExceptions(() => ShowNeighborSample(1)));
        TourTitle = toolbar.AddText(FontWeight.SemiBold);
        toolbar.EndGroup();
        TourGroup.IsVisible = false;

        LayoutRoot.Children.Add(toolbar);
        DockPanel.SetDock(toolbar, Dock.Top);
    }

    void InitializeCommands()
    {
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleGrid, first: true);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleOrtho);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleSnapToGrid);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleSnapToPoint);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleSnapToCenter);
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleLabelNewPoints);
        DrawingHost.AddToolbarButton(DrawingHost.CommandTogglePolar);
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

    GalleryView Gallery;
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

        Gallery.CanContinueDrawing = OwnDrawing != null;
        Gallery.IsVisible = true;
        LayoutRoot.IsVisible = false;
        UpdateTour();
        Publish(GalleryPath, AppTitle, push);
    }

    void ShowEditor()
    {
        Gallery.IsVisible = false;
        LayoutRoot.IsVisible = true;

        // the canvas must know its size before a drawing is fitted into it
        UpdateLayout();
        DrawingHost.DrawingControl.Focus();
    }

    void ShowNewDrawing(bool push)
    {
        ShowEditor();
        CurrentSample = null;
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

        ShowEditor();
        CurrentSample = null;
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
        ShowEditor();
        var control = DrawingHost.DrawingControl;
        if (OwnDrawing == null && CurrentSample == null && control.Drawing != null && control.Drawing.Figures.Any(figure => !(figure is CartesianGrid)))
        {
            OwnDrawing = control.Drawing;
        }

        CurrentSample = item;
        control.LoadDrawing(item.LoadText(), item.FileName);
        GalleryDrawing.Fit(control.Drawing, item.Plane);
        UpdateTour();
        Publish(item.Path, item.Title + " - " + AppTitle, push);
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
            TourPosition.Text = (GalleryCatalog.IndexOf(CurrentSample) + 1) + " / " + GalleryCatalog.Items.Count;
            TourTitle.Text = CurrentSample.Title;
        }
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
        ShowEditor();
        BecomeOwnDrawing();

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
        if (Gallery.IsVisible)
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
        if (Gallery.IsVisible)
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
