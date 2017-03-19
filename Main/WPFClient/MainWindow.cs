using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Printing;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using GuiLabs.Undo;
using System.Threading;

namespace DynamicGeometry
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            PageSettings = new Settings(this);
            AddBehaviors();
            LayoutRoot.Children.Add(DrawingHost);
            InitializeCommands();
            ViewDemoButton.IsEnabled = false;
            ThreadPool.QueueUserWorkItem(_ => DownloadDemoFile());
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

        /// <summary>
        /// Application Entry Point.
        /// </summary>
        [System.STAThreadAttribute()]
        public static void Main()
        {
            Application app = new Application();
            app.Run(new MainWindow());
        }

        private void InitializeComponent()
        {
            Title = "Live Geometry";
            WindowState = WindowState.Maximized;
            this.Content = LayoutRoot;
            //Icon = new BitmapImage(new Uri("/DG;component/Resources/Icons/DG.ico"));

            Menu menu = new Menu();
            LayoutRoot.Children.Add(menu);
            DockPanel.SetDock(menu, Dock.Top);

            MenuItem file = new MenuItem() { Header = "File" };
            MenuItem edit = new MenuItem() { Header = "Edit" };
            MenuItem view = new MenuItem() { Header = "View" };

            menu.Items.Add(file);
            menu.Items.Add(edit);
            menu.Items.Add(view);

            ClearButton = new MenuItem() { Header = "_New", InputGestureText = "Ctrl+N" };
            ClearButton.Click += ClearButton_Click;
            var open = new MenuItem() { Header = "_Open", InputGestureText = "Ctrl+O" };
            open.Click += Open_Click;
            var save = new MenuItem() { Header = "_Save...", InputGestureText = "Ctrl+S" };
            save.Click += Save_Click;
            var print = new MenuItem() { Header = "Print" };
            print.Click += Print_Click;
            var exit = new MenuItem() { Header = "Exit" };
            exit.Click += Exit_Click;

            var items = new UIElement[] {
                ClearButton,
                new Separator(),
                open,
                save,
                new Separator(),
                print,
                new Separator(),
                exit
            };

            file.ItemsSource = items;

            UndoButton = new MenuItem() { Header = "Undo", InputGestureText = "Ctrl+Z" };
            UndoButton.Click += Undo_Click;
            RedoButton = new MenuItem() { Header = "Redo", InputGestureText = "Ctrl+Y" };
            RedoButton.Click += Redo_Click;
            var cut = new MenuItem() { Header = "Cut", InputGestureText = "Ctrl+X", IsEnabled = false };
            cut.Click += Cut_Click;
            var copy = new MenuItem() { Header = "Copy", InputGestureText = "Ctrl+C" };
            copy.Click += Copy_Click;
            var paste = new MenuItem() { Header = "Paste", InputGestureText = "Ctrl+V" };
            paste.Click += Paste_Click;
            var pasteFrom = new MenuItem() { Header = "Paste from ..." };
            pasteFrom.Click += PasteFrom_Click;
            var delete = new MenuItem() { Header = "Delete", InputGestureText = "Del" };
            delete.Click += Delete_Click;
            var lockItem = new MenuItem() { Header = "Lock" };
            lockItem.Click += Lock_Click;
            var selectAll = new MenuItem() { Header = "Select all", InputGestureText = "Ctrl+A" };
            selectAll.Click += SelectAll_Click;
            var clear = new MenuItem() { Header = "Clear" };
            clear.Click += Clear_Click;

            items = new UIElement[] {
                UndoButton,
                RedoButton,
                new Separator(),
                cut,
                copy,
                paste,
                pasteFrom,
                delete,
                lockItem,
                new Separator(),
                selectAll,
                clear
            };

            edit.ItemsSource = items;

            ViewDemoButton = new MenuItem() { Header = "Demo" };
            ViewDemoButton.Click += ViewDemoButton_Click;
            ViewLibrary = new MenuItem() { Header = "Library", IsEnabled = false };
            SettingsButton = new MenuItem() { Header = "Settings" };
            SettingsButton.Click += SettingsButton_Click;
            FigureListButton = new MenuItem() { Header = "Figure List" };
            FigureListButton.Click += FigureListButton_Click;

            items = new UIElement[] {
                ViewDemoButton,
                new Separator(),
                ViewLibrary,
                new Separator(),
                SettingsButton,
                new Separator(),
                FigureListButton
            };

            view.ItemsSource = items;
        }

        MenuItem UndoButton;
        MenuItem RedoButton;
        MenuItem ClearButton;
        MenuItem ViewDemoButton;
        MenuItem ViewLibrary;
        MenuItem SettingsButton;
        MenuItem FigureListButton;

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

        DockPanel LayoutRoot = new DockPanel();
        DrawingHost DrawingHost = new DrawingHost();

        #region Demo

        private void DownloadDemoFile()
        {
            WebClient internet = new WebClient();
            internet.DownloadStringCompleted += new DownloadStringCompletedEventHandler(internet_DownloadStringCompleted);
            internet.DownloadStringAsync(new Uri("http://guilabs.de/geometry/demo/Demo.xml"));
        }

        /// <summary>
        /// http://www.csharp411.com/how-and-65279-and-other-byte-order-marks-bom-can-mess-up-your-xml/
        /// </summary>
        /// <remarks>
        /// Works as expected in Silverlight, but includes the Byte Order Mark in WPF
        /// </remarks>
        void internet_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            if (e.Cancelled || e.Error != null)
            {
                return;
            }
            demoText = e.Result;
#if !SILVERLIGHT
            // on WPF, need to strip the byte order mark for whatever reason (likely a bug)
            demoText = Utilities.StripByteOrderMark(demoText);
#endif
            ViewDemoButton.Dispatcher.BeginInvoke(new Action(() => ViewDemoButton.IsEnabled = true), DispatcherPriority.Background);
        }

        string demoText;

        private void ViewDemoButton_Click(object sender, RoutedEventArgs e)
        {
            DrawingHost.DrawingControl.LoadDrawing(demoText);
        }

        #endregion

        private void Undo_Click(object sender, RoutedEventArgs e)
        {
            DrawingHost.DrawingControl.Undo();
        }

        private void Redo_Click(object sender, RoutedEventArgs e)
        {
            DrawingHost.DrawingControl.Redo();
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.Close();
        }

        public void Load(string path)
        {
            if (path != null && File.Exists(path))
            {
                switch (System.IO.Path.GetExtension(path))
                {
                    case "." + dxfExtension:
                        HandleExceptions(() =>
                            DrawingHost.DrawingControl.Drawing = (new DXFDrawingDeserializer().ReadDrawing(path, DrawingHost.DrawingControl))
                        );

                        break;
                    default:
                        HandleExceptions(() =>
                            DrawingHost.DrawingControl.Drawing = Drawing.Load(path, DrawingHost.DrawingControl)
                        );

                        break;
                }
            }
        }

        public void Save(string path)
        {
            if (path != null)
            {
                HandleExceptions(() =>
                    DrawingHost.CurrentDrawing.Save(path)
                );
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            HandleExceptions(() =>
                DrawingHost.Clear()
            );
        }

        public void HandleExceptions(Action code)
        {
            try
            {
                code();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Open_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = drawingFileFilter;

            if (openFileDialog.ShowDialog().Value)
            {
                this.Load(openFileDialog.FileName);
            }
        }

        const string extension = "lgf";
        const string dxfExtension = "dxf";
        const string lgfFileFilter = "Live Geometry file (*." + extension + ")|*." + extension;
        const string pngFileFilter = "PNG image (*.png)|*.png";
        const string bmpFileFilter = "BMP image (*.bmp)|*.bmp";
        const string dxfFileFilter = "DXF file (*.dxf)|*.dxf";
        const string allFileFilter = "All files (*.*)|*.*";
        const string drawingFileFilter = lgfFileFilter + "|" + dxfFileFilter + "|" + allFileFilter;
        const string fileFilter = lgfFileFilter
                          + "|" + pngFileFilter
                          + "|" + bmpFileFilter
                          + "|" + allFileFilter;

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            HandleExceptions(() =>
            {
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                saveFileDialog.Filter = fileFilter;
                saveFileDialog.AddExtension = true;

                if (saveFileDialog.ShowDialog() != true)
                {
                    return;
                }

                var fileName = saveFileDialog.SafeFileName;
                var actualExtension = fileName.Substring(fileName.LastIndexOf('.')).ToLower();
                if (actualExtension == ".png")
                {
                    SaveToPng(DrawingHost.DrawingControl, saveFileDialog.FileName);
                }
                else if (actualExtension == ".bmp")
                {
                    SaveToBmp(DrawingHost.DrawingControl, saveFileDialog.FileName);
                }
                else
                {
                    Save(saveFileDialog.FileName);
                }
            }
            );
        }

        void SaveToBmp(FrameworkElement visual, string fileName)
        {
            var encoder = new BmpBitmapEncoder();
            SaveUsingEncoder(visual, fileName, encoder);
        }

        void SaveToPng(FrameworkElement visual, string fileName)
        {
            var encoder = new PngBitmapEncoder();
            SaveUsingEncoder(visual, fileName, encoder);
        }

        void SaveUsingEncoder(FrameworkElement visual, string fileName, BitmapEncoder encoder)
        {
            RenderTargetBitmap bitmap = new RenderTargetBitmap(
                (int)visual.ActualWidth,
                (int)visual.ActualHeight,
                96,
                96,
                PixelFormats.Pbgra32);
            bitmap.Render(visual);
            BitmapFrame frame = BitmapFrame.Create(bitmap);
            encoder.Frames.Add(frame);

            using (var stream = File.Create(fileName))
            {
                encoder.Save(stream);
            }
        }

        private void Cut_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            HandleExceptions(() =>
                DrawingHost.CurrentDrawing.Copy()
            );
        }

        private void Paste_Click(object sender, RoutedEventArgs e)
        {
            HandleExceptions(() =>
                DrawingHost.CurrentDrawing.Paste()
            );
        }

        private void PasteFrom_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = drawingFileFilter;

            if (openFileDialog.ShowDialog().Value)
            {
                HandleExceptions(() =>
                    DrawingHost.CurrentDrawing.PasteFrom(openFileDialog.FileName)
                );
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            DeleteSelection();
        }

        private void DeleteSelection()
        {
            HandleExceptions(() =>
                DrawingHost.CurrentDrawing.DeleteSelection()
            );
        }

        private void LockSelection()
        {
            HandleExceptions(() =>
                DrawingHost.CurrentDrawing.LockSelected()
            );
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            SelectAll();
        }

        private void FigureListButton_Click(object sender, RoutedEventArgs e)
        {
            HandleExceptions(() =>
                DrawingHost.CommandShowFigureExplorer.Execute()
            );
        }

        private void SelectAll()
        {
            HandleExceptions(() =>
                DrawingHost.CurrentDrawing.SelectAll()
            );
        }

        private void Lock_Click(object sender, RoutedEventArgs e)
        {
            HandleExceptions(() =>
                DrawingHost.CurrentDrawing.LockSelected()
            );
        }

        protected override void OnKeyUp(System.Windows.Input.KeyEventArgs e)
        {
            base.OnKeyUp(e);

            switch (e.Key)
            {
                case System.Windows.Input.Key.Z:
                    if (e.KeyboardDevice.IsKeyDown(System.Windows.Input.Key.RightCtrl) || e.KeyboardDevice.IsKeyDown(System.Windows.Input.Key.LeftCtrl))
                        DrawingHost.DrawingControl.Undo();
                    break;
                case System.Windows.Input.Key.Y:
                    if (e.KeyboardDevice.IsKeyDown(System.Windows.Input.Key.RightCtrl) || e.KeyboardDevice.IsKeyDown(System.Windows.Input.Key.LeftCtrl))
                        DrawingHost.DrawingControl.Redo();
                    break;
                case System.Windows.Input.Key.A:
                    if (e.KeyboardDevice.IsKeyDown(System.Windows.Input.Key.RightCtrl) || e.KeyboardDevice.IsKeyDown(System.Windows.Input.Key.LeftCtrl))
                        SelectAll();
                    break;
                case System.Windows.Input.Key.Delete:
                    DeleteSelection();
                    break;
                case System.Windows.Input.Key.Escape:
                    DrawingHost.CurrentDrawing.Behavior.Restart();
                    break;
            }
        }

        #region Geometry Toolbar

        private bool mShowToolbarButtonText = false;
        public bool ShowToolbarButtonText
        {
            get
            {
                return mShowToolbarButtonText;
            }
            set
            {
                mShowToolbarButtonText = value;
            }
        }

        private bool mShowCoordinateGrid = false;
        public bool ShowCoordinateGrid
        {
            get
            {
                return mShowCoordinateGrid;
            }
            set
            {
                mShowCoordinateGrid = value;
                DrawingHost.CurrentDrawing.CoordinateGrid.Visible = value;
            }
        }

        private void ShowToolbarText_Checked(object sender, RoutedEventArgs e)
        {
            ShowToolbarButtonText = !ShowToolbarButtonText;
        }

        #endregion

        #region Undo/Redo

        private void UndoButton_Click(object sender, RoutedEventArgs e)
        {
            DrawingHost.DrawingControl.Undo();
            UpdateUndoRedo();
            DrawingHost.ShowProperties(null);
        }

        private void RedoButton_Click(object sender, RoutedEventArgs e)
        {
            DrawingHost.DrawingControl.Redo();
            UpdateUndoRedo();
            DrawingHost.ShowProperties(null);
        }

        private void UpdateUndoRedo()
        {
            UndoButton.IsEnabled = DrawingHost.CurrentDrawing.ActionManager.CanUndo; // ? Visibility.Visible : Visibility.Collapsed;
            RedoButton.IsEnabled = DrawingHost.CurrentDrawing.ActionManager.CanRedo; // ? Visibility.Visible : Visibility.Collapsed;
            ClearButton.IsEnabled = DrawingHost.CurrentDrawing.Figures.Count > 0;
            DrawingHost.CurrentDrawing.ClearStatus();
        }

        void ActionManager_CollectionChanged(object sender, EventArgs e)
        {
            UpdateUndoRedo();
        }

        #endregion

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            DrawingHost.Clear();
        }

        #region Settings

        Settings PageSettings;

        public class Settings
        {
            public Settings(MainWindow page)
            {
                Page = page;
                this.DebugInfo = new DebugInformation(page);
            }

            MainWindow Page;

            [PropertyGridVisible]
            [PropertyGridName("Show coordinate axes and grid")]
            public bool ShowGrid
            {
                get
                {
                    return Page.ShowCoordinateGrid;
                }
                set
                {
                    Page.ShowCoordinateGrid = value;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Show text in toolbar buttons")]
            public bool ShowToolbarText
            {
                get
                {
                    return Page.ShowToolbarButtonText;
                }
                set
                {
                    Page.ShowToolbarButtonText = value;
                }
            }

            //[PropertyGridVisible]
            //[PropertyGridName("Toolbar orientation")]
            //public Orientation ToolbarOrientation
            //{
            //    get
            //    {
            //        return Page.ToolbarOrientation;
            //    }
            //    set
            //    {
            //        Page.ToolbarOrientation = value;
            //    }
            //}

            public class DebugInformation
            {
                public DebugInformation(MainWindow page)
                {
                    Page = page;
                }

                MainWindow Page;

                [PropertyGridVisible]
                public string DrawingXml
                {
                    get
                    {
                        return Page.DrawingHost.CurrentDrawing.SaveAsText();
                    }
                }

                [PropertyGridVisible]
                public IEnumerable<IAction> UndoBuffer
                {
                    get
                    {
                        return Page.DrawingHost.CurrentDrawing.ActionManager.EnumUndoableActions().ToArray();
                    }
                }
            }

            [PropertyGridVisible]
            public DebugInformation DebugInfo { get; set; }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
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

        Point findIntersection(Point p1, Point p2, Point p3, Point p4)
        {
            double xD1, yD1, xD2, yD2, xD3, yD3;
            double dot, deg, len1, len2;
            double segmentLen1, segmentLen2;
            double ua, ub, div;

            // calculate differences  
            xD1 = p2.x - p1.x;
            xD2 = p4.x - p3.x;
            yD1 = p2.y - p1.y;
            yD2 = p4.y - p3.y;
            xD3 = p1.x - p3.x;
            yD3 = p1.y - p3.y;

            // calculate the lengths of the two lines  
            len1 = Math.Sqr(xD1 * xD1 + yD1 * yD1);
            len2 = Math.Sqr(xD2 * xD2 + yD2 * yD2);

            // calculate angle between the two lines.  
            dot = (xD1 * xD2 + yD1 * yD2); // dot product  
            deg = dot / (len1 * len2);

            // if abs(angle)==1 then the lines are parallell,  
            // so no intersection is possible  
            if (Math.Abs(deg) == 1) return null;

            // find intersection Pt between two lines  
            Point pt = new Point(0, 0);
            div = yD2 * xD1 - xD2 * yD1;
            ua = (xD2 * yD3 - yD2 * xD3) / div;
            ub = (xD1 * yD3 - yD1 * xD3) / div;
            pt.x = p1.x + ua * xD1;
            pt.y = p1.y + ua * yD1;

            // calculate the combined length of the two segments  
            // between Pt-p1 and Pt-p2  
            xD1 = pt.x - p1.x;
            xD2 = pt.x - p2.x;
            yD1 = pt.y - p1.y;
            yD2 = pt.y - p2.y;
            segmentLen1 = Math.Sqr(xD1 * xD1 + yD1 * yD1) + Math.Sqr(xD2 * xD2 + yD2 * yD2);

            // calculate the combined length of the two segments  
            // between Pt-p3 and Pt-p4  
            xD1 = pt.x - p3.x;
            xD2 = pt.x - p4.x;
            yD1 = pt.y - p3.y;
            yD2 = pt.y - p4.y;
            segmentLen2 = Math.Sqr(xD1 * xD1 + yD1 * yD1) + Math.Sqr(xD2 * xD2 + yD2 * yD2);

            // if the lengths of both sets of segments are the same as  
            // the lenghts of the two lines the point is actually  
            // on the line segment.  

            // if the point isn’t on the line, return null  
            if (Math.Abs(len1 - segmentLen1) > 0.01 || Math.Abs(len2 - segmentLen2) > 0.01)
                return null;

            // return the valid intersection  
            return pt;
        }

        class Point
        {
            public double x, y;
            public Point(double x, double y)
            {
                this.x = x;
                this.y = y;
            }

            void set(double x, double y)
            {
                this.x = x;
                this.y = y;
            }
        }

        private void Print_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog dialog = new PrintDialog();
            if (dialog.ShowDialog() == true)
            {
                StackPanel panel = new StackPanel();
                panel.Margin = new Thickness(15);
                panel.Children.Add((Canvas)XamlReader.Parse(XamlWriter.Save(this.DrawingHost.DrawingControl)));
                TextBlock myBlock = new TextBlock();
                myBlock.Text = "Drawing";
                myBlock.TextAlignment = TextAlignment.Center;
                panel.Children.Add(myBlock);

                panel.Measure(new Size(dialog.PrintableAreaWidth,
                  dialog.PrintableAreaHeight));
                panel.Arrange(new Rect(new System.Windows.Point(0, 0), panel.DesiredSize));
                dialog.PrintTicket.PageOrientation = PageOrientation.Landscape;
                dialog.PrintVisual(panel, "Drawing");
            }
        }
         

    }
}
