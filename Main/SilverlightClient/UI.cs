using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using DynamicGeometry;
using ImageTools;
using ImageTools.IO;
using ImageTools.IO.Bmp;
using ImageTools.IO.Png;

namespace LiveGeometry
{
    public partial class Page
    {
        void Page_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            try
            {
                if (Behavior.IsCtrlPressed())
                {
                    if (e.Key == System.Windows.Input.Key.Z)
                    {
                        drawingHost.DrawingControl.CommandUndo.Execute();
                        e.Handled = true;
                        return;
                    }
                    else if (e.Key == Key.Y)
                    {
                        drawingHost.DrawingControl.CommandRedo.Execute();
                        e.Handled = true;
                        return;
                    }
                }

                if (e.Key == Key.F9)
                {
                    PageSettings.Fullscreen();
                    e.Handled = true;
                    return;
                }

                if (drawingHost != null && drawingHost.CurrentDrawing != null && !e.Handled)
                {
                    drawingHost.CurrentDrawing.Behavior.KeyDown(sender, e);
                }
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        public class ExceptionMessageDialog : MessageBoxDialog
        {
            public ExceptionMessageDialog(Page page, Exception ex, string message)
            {
                parent = page;
                details = new Details(ex.ToString());
                errorCode = ex.ToString().GetHashCode().ToString();
                MessageText = message;
            }

            Page parent;

            [PropertyGridVisible]
            public override string Message
            {
                get
                {
                    return base.Message;
                }
            }

            string errorCode;

            [PropertyGridVisible]
            public string Error
            {
                get
                {
                    return errorCode;
                }
            }

            Details details;
            [PropertyGridVisible]
            [PropertyGridName("Details")]
            public Details ErrorDetails
            {
                get
                {
                    return details;
                }
            }

            public class Details
            {
                public Details(string details)
                {
                    errorDetails = details;
                }

                string errorDetails;

                [PropertyGridVisible]
                [PropertyGridName("Stack trace:")]
                public string ErrorDetails
                {
                    get
                    {
                        return errorDetails;
                    }
                }
            }

            protected override void OKClicked()
            {
                parent.drawingHost.DrawingControl.Drawing.ClearStatus();
            }

            public override string ToString()
            {
                return "Error";
            }
        }

        void HandleException(Exception ex)
        {
            var message =
                "Live Geometry has just encountered an error.\n" +
                "The error details will be reported automatically \n" +
                "and the bug will be fixed as soon as possible.\n\n" +
                "No personal data is transmitted.\n\n" +
                "If you have any questions, please feel free to go to \n" +
                "http://livegeometry.codeplex.com/Thread/List.aspx \n" +
                "and mention the error code from below.\n\n" +
                "Thanks for making Live Geometry better!";

            var dialog = new ExceptionMessageDialog(this, ex, message);
            drawingHost.ShowProperties(dialog);
            if (liveGeometryWebServices == null)
            {
                return;
            }
            try
            {
                liveGeometryWebServices.SendErrorReportAsync(ex.ToString());
            }
            catch (Exception second)
            {
                drawingHost.ShowHint(second.ToString());
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            drawingHost.Clear();
        }

        const string extension = "lgf";
        const string lgfFileFilter = "Live Geometry file (*." + extension + ")|*." + extension;
        const string lgfdgfFileFilter = "Live Geometry file (*.lgf, *.dgf)|*.lgf;*.dgf";
        const string pngFileFilter = "PNG image (*.png)|*.png";
        const string bmpFileFilter = "BMP image (*.bmp)|*.bmp";
        const string dgfFileFilter = "DG Drawing (*.dgf)|*.dgf";
        const string allFileFilter = "All files (*.*)|*.*";
        const string fileFilter = lgfFileFilter
                          + "|" + pngFileFilter
                          + "|" + bmpFileFilter
                          + "|" + allFileFilter
                          ;
        const string openFileFilter = lgfdgfFileFilter
                          + "|" + allFileFilter
                          ;

        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            Open();
        }

        private void Open()
        {
            if (Application.Current.HasElevatedPermissions && drawingHost.CurrentDrawing.HasUnsavedChanges)
            {
                ShowUnsavedChangesDialog(ShowOpenFileDialog);
            }
            ShowOpenFileDialog();
        }

        private void ShowOpenFileDialog()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = openFileFilter;
            dialog.Multiselect = false;
            var result = dialog.ShowDialog();
            if (result != true)
            {
                return;
            }

            string text = null;

            var extension = dialog.File.Extension;
            if (extension.Equals(".dgf", StringComparison.OrdinalIgnoreCase))
            {
                using (var stream = dialog.File.OpenRead())
                {
                    int length = (int)stream.Length;
                    byte[] array = new byte[length];
                    stream.Read(array, 0, length);
                    text = Encoding.FromWindows1251(array);
                    string[] lines = text.Split(
                        new[] { "\r\n", "\r", "\n" },
                        StringSplitOptions.RemoveEmptyEntries);
                    if (lines.Length > 10)
                    {
                        drawingHost.DrawingControl.LoadDrawingFromDGF(lines);
                    }
                    else
                    {
                        drawingHost.ShowHint("Incorrect DGF file (it should contain more than 10 text lines)");
                    }
                }
                return;
            }

            using (var sr = dialog.File.OpenText())
            {
                text = sr.ReadToEnd();
            }
            if (!string.IsNullOrEmpty(text))
            {
                drawingHost.DrawingControl.LoadDrawing(text);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Save();
        }

        private void Save()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = fileFilter;
            dialog.FilterIndex = 1;
            dialog.DefaultExt = extension;

            var result = dialog.ShowDialog();
            if (result != true)
            {
                return;
            }

            try
            {
                var fileName = dialog.SafeFileName;
                var actualExtension = fileName.Substring(fileName.LastIndexOf('.')).ToLower();
                if (actualExtension == ".png")
                {
                    SaveAsPng(drawingHost.DrawingControl, dialog);
                }
                else if (actualExtension == ".bmp")
                {
                    SaveAsBmp(drawingHost.DrawingControl, dialog);
                }
                else
                {
                    SaveDrawingToLGF(dialog);
                }
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        private void New()
        {
            if (Application.Current.HasElevatedPermissions && drawingHost.CurrentDrawing.HasUnsavedChanges)
            {
                ShowUnsavedChangesDialog(drawingHost.Clear);
            }
            else
            {
                drawingHost.Clear();
            }
        }

        public void ShowUnsavedChangesDialog(Action subsequentAction)
        {
            var saveButton = new Button()
            {
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(8),
                Content = "Save",
                Width = 100
            };
            saveButton.Click += delegate(object sender, RoutedEventArgs e)
            {
                (((sender as Button).Parent as Grid).Parent as ChildWindow).Close();
                Save();
                subsequentAction();
            };

            var dontSaveButton = new Button()
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(8),
                Content = "Don't Save",
                Width = 100
            };
            dontSaveButton.Click += delegate(object sender, RoutedEventArgs e)
            {
                (((sender as Button).Parent as Grid).Parent as ChildWindow).Close();
                subsequentAction();
            };

            var cancelButton = new Button()
            {
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(8),
                Content = "Cancel",
                Width = 100
            };
            cancelButton.Click += delegate(object sender, RoutedEventArgs e)
            {
                (((sender as Button).Parent as Grid).Parent as ChildWindow).Close();
            };

            TextBlock text = new TextBlock();
            string drawingName = drawingHost.CurrentDrawing.Name;
            if (drawingName == null) drawingName = "untitled";
            text.Text = "Do you want to save changes to \"" + drawingName + "\"?";

            Grid grid = new Grid()
            {
                Width = 360
            };
            grid.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
            Grid.SetRow(text, 0);
            Grid.SetRow(dontSaveButton, 1);
            Grid.SetRow(saveButton, 1);
            Grid.SetRow(cancelButton, 1);
            grid.Children.Add(text, saveButton, dontSaveButton, cancelButton);

            var unsavedChangesWindow = new ChildWindow()
            {
                Content = grid
            };
            unsavedChangesWindow.HasCloseButton = false;
            unsavedChangesWindow.Show();
        }

        void SaveAsPng(Canvas canvas, SaveFileDialog dialog)
        {
            SaveToImage(canvas, dialog, new PngEncoder());
        }

        void SaveAsBmp(Canvas canvas, SaveFileDialog dialog)
        {
            SaveToImage(canvas, dialog, new BmpEncoder());
        }

        void SaveToImage(Canvas canvas, SaveFileDialog dialog, IImageEncoder encoder)
        {
            using (var stream = dialog.OpenFile())
            {
                var image = canvas.ToImage();
                encoder.Encode(image, stream);
            }
        }

        void SaveDrawingToLGF(SaveFileDialog dialog)
        {
            var currentDrawing = drawingHost.CurrentDrawing;
            DrawingSerializer.SaveDrawing(currentDrawing, dialog.OpenFile());
            if (Application.Current.HasElevatedPermissions)
            {
                var undoableActions = currentDrawing.ActionManager.EnumUndoableActions();
                if (!undoableActions.IsEmpty())
                {
                    currentDrawing.LastUndoableActionAtSave = undoableActions.Last();
                }
            }
        }
    }
}
