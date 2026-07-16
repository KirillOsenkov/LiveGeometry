using System;
using System.IO;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using DynamicGeometry;

namespace LiveGeometry;

/// <summary>
/// The shared Live Geometry main view, used by both the Browser and Desktop heads.
/// Ported from Main/WPFClient/MainWindow.cs.
/// </summary>
public class MainView : UserControl
{
    DockPanel LayoutRoot = new DockPanel();
    DrawingHost DrawingHost = new DrawingHost();

    MenuItem UndoButton;
    MenuItem RedoButton;
    MenuItem ClearButton;

    static readonly FilePickerFileType LgfFileType = new("Live Geometry drawing")
    {
        Patterns = new[] { "*.lgf" }
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


        AddHandler(KeyUpEvent, MainView_KeyUp, RoutingStrategies.Tunnel);

        AttachedToVisualTree += (s, e) =>
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel != null)
            {
                DynamicGeometry.Clipboard.SystemClipboard = topLevel.Clipboard;
            }

            Focus();
        };
    }

    private void AddBehaviors()
    {
        var behaviors = Behavior.LoadBehaviors(typeof(Dragger).Assembly);
        Behavior.Default = behaviors.First(b => b is Dragger);
        foreach (var behavior in behaviors)
        {
            DrawingHost.AddToolButton(behavior);
        }
    }

    private void InitializeComponent()
    {
        Focusable = true;
        Content = LayoutRoot;

        Menu menu = new Menu();
        LayoutRoot.Children.Add(menu);
        DockPanel.SetDock(menu, Dock.Top);

        MenuItem file = new MenuItem() { Header = "File" };
        MenuItem edit = new MenuItem() { Header = "Edit" };
        MenuItem view = new MenuItem() { Header = "View" };

        menu.Items.Add(file);
        menu.Items.Add(edit);
        menu.Items.Add(view);

        ClearButton = AddItem(file, "_New", ClearButton_Click, new KeyGesture(Key.N, KeyModifiers.Control));
        file.Items.Add(new Separator());
        AddItem(file, "_Open…", Open_Click, new KeyGesture(Key.O, KeyModifiers.Control));
        AddItem(file, "_Save…", Save_Click, new KeyGesture(Key.S, KeyModifiers.Control));

        UndoButton = AddItem(edit, "Undo", Undo_Click, new KeyGesture(Key.Z, KeyModifiers.Control));
        RedoButton = AddItem(edit, "Redo", Redo_Click, new KeyGesture(Key.Y, KeyModifiers.Control));
        edit.Items.Add(new Separator());
        AddItem(edit, "Copy", Copy_Click, new KeyGesture(Key.C, KeyModifiers.Control));
        AddItem(edit, "Paste", Paste_Click, new KeyGesture(Key.V, KeyModifiers.Control));
        AddItem(edit, "Delete", Delete_Click, new KeyGesture(Key.Delete));
        AddItem(edit, "Lock", Lock_Click, null);
        edit.Items.Add(new Separator());
        AddItem(edit, "Select all", SelectAll_Click, new KeyGesture(Key.A, KeyModifiers.Control));
        AddItem(edit, "Clear", Clear_Click, null);

        AddItem(view, "Settings", SettingsButton_Click, null);
        AddItem(view, "Figure List", FigureListButton_Click, null);
    }

    MenuItem AddItem(MenuItem parent, string header, EventHandler<RoutedEventArgs> onClick, KeyGesture gesture)
    {
        var item = new MenuItem() { Header = header };
        if (gesture != null)
        {
            item.InputGesture = gesture;
        }
        item.Click += (s, e) => onClick(s, e);
        parent.Items.Add(item);
        return item;
    }

    void InitializeCommands()
    {
        DrawingHost.AddToolbarButton(DrawingHost.CommandToggleGrid);
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

    private void Undo_Click(object sender, RoutedEventArgs e) => DrawingHost.DrawingControl.Undo();

    private void Redo_Click(object sender, RoutedEventArgs e) => DrawingHost.DrawingControl.Redo();

    private void ClearButton_Click(object sender, RoutedEventArgs e) => DrawingHost.Clear();

    private void Clear_Click(object sender, RoutedEventArgs e) => HandleExceptions(() => DrawingHost.Clear());

    private async void Open_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var topLevel = TopLevel.GetTopLevel(this);
            var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Open drawing",
                FileTypeFilter = new[] { LgfFileType, FilePickerFileTypes.All }
            });

            var file = files?.FirstOrDefault();
            if (file == null)
            {
                return;
            }

            string text;
            using (var stream = await file.OpenReadAsync())
            using (var reader = new StreamReader(stream))
            {
                text = await reader.ReadToEndAsync();
            }

            text = Utilities.StripByteOrderMark(text);
            HandleExceptions(() => DrawingHost.DrawingControl.LoadDrawing(text, file.Name));
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    private async void Save_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var topLevel = TopLevel.GetTopLevel(this);
            var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Save drawing",
                SuggestedFileName = "drawing.lgf",
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
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    private void Copy_Click(object sender, RoutedEventArgs e) => HandleExceptions(() => DrawingHost.CurrentDrawing.Copy());

    private void Paste_Click(object sender, RoutedEventArgs e) => HandleExceptions(() => DrawingHost.CurrentDrawing.Paste());

    private void Delete_Click(object sender, RoutedEventArgs e) => DeleteSelection();

    private void DeleteSelection() => HandleExceptions(() => DrawingHost.CurrentDrawing.DeleteSelection());

    private void Lock_Click(object sender, RoutedEventArgs e) => HandleExceptions(() => DrawingHost.CurrentDrawing.LockSelected());

    private void SelectAll_Click(object sender, RoutedEventArgs e) => SelectAll();

    private void SelectAll() => HandleExceptions(() => DrawingHost.CurrentDrawing.SelectAll());

    private void FigureListButton_Click(object sender, RoutedEventArgs e) =>
        HandleExceptions(() => DrawingHost.CommandShowFigureExplorer.Execute());

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

    private void MainView_KeyUp(object sender, KeyEventArgs e)
    {
        var focused = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement();
        if (focused is TextBox)
        {
            return;
        }

        bool ctrl = (e.KeyModifiers & KeyModifiers.Control) != 0;
        switch (e.Key)
        {
            case Key.Z:
                if (ctrl)
                {
                    DrawingHost.DrawingControl.Undo();
                }
                break;
            case Key.Y:
                if (ctrl)
                {
                    DrawingHost.DrawingControl.Redo();
                }
                break;
            case Key.A:
                if (ctrl)
                {
                    SelectAll();
                }
                break;
            case Key.Delete:
                DeleteSelection();
                break;
            case Key.Escape:
                if (DrawingHost.CurrentDrawing.Behavior.IsInInitialState)
                {
                    DrawingHost.CurrentDrawing.SetDefaultBehavior();
                }
                else
                {
                    DrawingHost.CurrentDrawing.Behavior.Restart();
                }

                e.Handled = true;
                break;
        }
    }
}
