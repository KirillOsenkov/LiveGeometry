using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;
using SilverlightContrib.Controls;

namespace GuiLabs.Controls
{
    /// <summary>
    /// Represents a Color Picker control which allows a user to select a color.
    /// </summary>
    public class ColorPicker : Grid
    {
        private Grid SelectedColorSample;
        private Grid TransparencyField;
        private Grid ColorField;
        private Grid ColorSelector;
        private Grid HueSelector;
        private Grid AlphaSelector;
        private Grid ChooserArea;
        private Rectangle ColorSample;
        private Rectangle TransparencyGradient;
        private Canvas Rainbow;

        public ColorPicker()
        {
            this.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
            this.RowDefinitions.Add(new RowDefinition());
            this.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
            this.Width = 200;

            SelectedColorSample = CreateColorSample();
            Rainbow = CreateRainbow();
            ColorField = CreateColorField();
            TransparencyField = CreateTransparencyField();

            ChooserArea = new Grid();
            ChooserArea.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
            ChooserArea.ColumnDefinitions.Add(new ColumnDefinition());
            ChooserArea.Children.Add(Rainbow);
            ChooserArea.Children.Add(ColorField);
            Grid.SetColumn(ColorField, 1);

            this.Children.Add(SelectedColorSample);
            this.Children.Add(ChooserArea);
            this.Children.Add(TransparencyField);

            Grid.SetRow(ChooserArea, 1);
            Grid.SetRow(TransparencyField, 2);

            Expanded = false;
            UpdateSelectedColorSample();
        }

        private Grid CreateColorSample()
        {
            var result = new Grid();
            result.Background = Brushes.White;

            var canvas = CreateCheckerboard();
            result.Children.Add(canvas);

            ColorSample = new Rectangle();
            result.MinHeight = 20;
            result.HorizontalAlignment = HorizontalAlignment.Stretch;
            result.VerticalAlignment = VerticalAlignment.Stretch;
            result.MouseLeftButtonDown += ColorSample_MouseLeftButtonDown;
            ColorSample.Fill = new SolidColorBrush(Colors.Black);
            ColorSample.HorizontalAlignment = HorizontalAlignment.Stretch;
            ColorSample.VerticalAlignment = VerticalAlignment.Stretch;

            result.Children.Add(ColorSample);

            return result;
        }

        private Grid CreateTransparencyField()
        {
            var result = new Grid();
            result.Background = Brushes.White;

            var canvas = CreateCheckerboard();
            canvas.HorizontalAlignment = HorizontalAlignment.Stretch;
            result.Children.Add(canvas);

            TransparencyGradient = new Rectangle()
            {
                IsHitTestVisible = false,
                Width = 200,
                Height = 20
            };
            canvas.Children.Add(TransparencyGradient);

            result.HorizontalAlignment = HorizontalAlignment.Stretch;
            result.VerticalAlignment = VerticalAlignment.Stretch;
            result.MinHeight = 20;
            result.MouseLeftButtonDown += TransparencyField_MouseLeftButtonDown;
            result.MouseMove += TransparencyField_MouseMove;
            result.MouseLeftButtonUp += TransparencyField_MouseLeftButtonUp;

            AlphaSelector = CreateAlphaSelector();
            canvas.Children.Add(AlphaSelector);

            return result;
        }

        private Canvas CreateCheckerboard()
        {
            var canvas = new Canvas();
            for (int i = 0; i < 10; i++)
            {
                AddBlackRectangle(canvas, i * 20, 0);
                AddBlackRectangle(canvas, i * 20 + 10, 10);
            }
            return canvas;
        }

        private void AddBlackRectangle(Canvas result, int x, int y)
        {
            var rect = new Rectangle()
            {
                Width = 10,
                Height = 10,
                Fill = new SolidColorBrush(Colors.LightGray),
                IsHitTestVisible = false
            };

            Canvas.SetLeft(rect, x);
            Canvas.SetTop(rect, y);
            result.Children.Add(rect);
        }

        private Grid CreateAlphaSelector()
        {
            var result = new Grid();
            var triangle1 = new Polygon()
            {
                Points = new PointCollection()
                {
                    new Point(0, 0),
                    new Point(5, 10),
                    new Point(10, 0)
                },
                Fill = new SolidColorBrush(Colors.Black),
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top
            };
            var triangle2 = new Polygon()
            {
                Points = new PointCollection()
                {
                    new Point(0, 10),
                    new Point(5, 0),
                    new Point(10, 10)
                },
                Fill = new SolidColorBrush(Colors.Black),
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Bottom
            };

            result.Children.Add(triangle1);
            result.Children.Add(triangle2);
            result.VerticalAlignment = VerticalAlignment.Top;
            result.HorizontalAlignment = HorizontalAlignment.Left;
            result.IsHitTestVisible = false;
            return result;
        }

        private Canvas CreateRainbow()
        {
            HueSelector = CreateHueSelector();

            var result = new Canvas();
            result.Background = new LinearGradientBrush(
                new GradientStopCollection()
                {
                    new GradientStop() { Offset = 0.00, Color = Color.FromArgb(255, 255, 0, 0)},
                    new GradientStop() { Offset = 0.17, Color = Color.FromArgb(255, 255, 255, 0) },
                    new GradientStop() { Offset = 0.33, Color = Color.FromArgb(255, 0, 255, 0) },
                    new GradientStop() { Offset = 0.50, Color = Color.FromArgb(255, 0, 255, 255) },
                    new GradientStop() { Offset = 0.66, Color = Color.FromArgb(255, 0, 0, 255) },
                    new GradientStop() { Offset = 0.83, Color = Color.FromArgb(255, 255, 0, 255) },
                    new GradientStop() { Offset = 1.00, Color = Color.FromArgb(255, 255, 0, 0) },
                },
                90);
            result.HorizontalAlignment = HorizontalAlignment.Stretch;
            result.VerticalAlignment = VerticalAlignment.Stretch;
            result.MinWidth = 20;
            result.MouseLeftButtonDown += Rainbow_MouseLeftButtonDown;
            result.MouseMove += Rainbow_MouseMove;
            result.MouseLeftButtonUp += Rainbow_MouseLeftButtonUp;
            result.Children.Add(HueSelector);
            return result;
        }

        private Grid CreateHueSelector()
        {
            var result = new Grid();
            var triangle1 = new Polygon()
            {
                Points = new PointCollection()
                {
                    new Point(0, 0),
                    new Point(10, 5),
                    new Point(0, 10)
                },
                Fill = new SolidColorBrush(Colors.Black),
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top
            };
            var triangle2 = new Polygon()
            {
                Points = new PointCollection()
                {
                    new Point(10, 0),
                    new Point(0, 5),
                    new Point(10, 10)
                },
                Fill = new SolidColorBrush(Colors.Black),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top
            };

            result.Children.Add(triangle1);
            result.Children.Add(triangle2);
            result.VerticalAlignment = VerticalAlignment.Top;
            result.HorizontalAlignment = HorizontalAlignment.Left;
            result.IsHitTestVisible = false;
            return result;
        }

        private Grid CreateColorField()
        {
            var whiteGradient = new Rectangle()
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
            };
            var blackGradient = new Rectangle()
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            var canvas = new Canvas()
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            ColorSelector = CreateColorSelector();
            canvas.Children.Add(ColorSelector);

            whiteGradient.Fill = new LinearGradientBrush(
                new GradientStopCollection()
                {
                    new GradientStop() { Color = Color.FromArgb(255, 255, 255, 255), Offset = 0 },
                    new GradientStop() { Color = Color.FromArgb(0, 255, 255, 255), Offset = 1 }
                },
                0);

            blackGradient.Fill = new LinearGradientBrush(
                new GradientStopCollection()
                {
                    new GradientStop() { Color = Color.FromArgb(0, 0, 0, 0), Offset = 0 },
                    new GradientStop() { Color = Color.FromArgb(255, 0, 0, 0), Offset = 1 },
                },
                90);

            var result = new Grid();
            result.Background = new SolidColorBrush(Colors.Red);
            result.HorizontalAlignment = HorizontalAlignment.Stretch;
            result.VerticalAlignment = VerticalAlignment.Stretch;
            result.MinWidth = 180;
            result.MinHeight = result.MinWidth;
            result.Children.Add(whiteGradient);
            result.Children.Add(blackGradient);
            result.Children.Add(canvas);
            result.MouseLeftButtonDown += ColorField_MouseLeftButtonDown;
            result.MouseMove += ColorField_MouseMove;
            result.MouseLeftButtonUp += ColorField_MouseLeftButtonUp;
            return result;
        }

        private Grid CreateColorSelector()
        {
            var result = new Grid()
            {
                IsHitTestVisible = false
            };
            var ellipse1 = new Ellipse()
            {
                Width = 10,
                Height = 10,
                StrokeThickness = 3,
                Stroke = new SolidColorBrush(Colors.White)
            };
            var ellipse2 = new Ellipse()
            {
                Width = 10,
                Height = 10,
                StrokeThickness = 1,
                Stroke = new SolidColorBrush(Colors.Black)
            };
            result.Children.Add(ellipse1);
            result.Children.Add(ellipse2);
            result.HorizontalAlignment = HorizontalAlignment.Left;
            result.VerticalAlignment = VerticalAlignment.Top;
            result.IsHitTestVisible = false;
            return result;
        }

        private void ColorSample_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Expanded = !Expanded;
        }

        public bool Expanded
        {
            get
            {
                return ChooserArea.Visibility == Visibility.Visible;
            }
            set
            {
                var visibility = value ? Visibility.Visible : Visibility.Collapsed;
                if (ChooserArea.Visibility == visibility)
                {
                    return;
                }
                ChooserArea.Visibility = visibility;
                TransparencyField.Visibility = visibility;
                if (value)
                {
#if SILVERLIGHT
                    Dispatcher.BeginInvoke(() => SelectedColor = SelectedColor);
#else
                    Dispatcher.BeginInvoke(new System.Action(() => SelectedColor = SelectedColor), DispatcherPriority.Render);
#endif
                }
            }
        }

        private void Rainbow_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            RainbowMouseCaptured = Rainbow.CaptureMouse();
            UpdateHuePos(e.GetPosition(Rainbow).Y);
        }

        private void Rainbow_MouseMove(object sender, MouseEventArgs e)
        {
            if (ColorFieldMouseCaptured || TransparencyFieldMouseCaptured || !RainbowMouseCaptured)
            {
                return;
            }

            UpdateHuePos(e.GetPosition(Rainbow).Y);
        }

        private void Rainbow_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            Rainbow.ReleaseMouseCapture();
            RainbowMouseCaptured = false;
            SelectedColor = GetColor();
        }

        private void TransparencyField_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            TransparencyFieldMouseCaptured = TransparencyField.CaptureMouse();
            UpdateAlpha(e.GetPosition(TransparencyField).X);
        }

        private void TransparencyField_MouseMove(object sender, MouseEventArgs e)
        {
            if (ColorFieldMouseCaptured || RainbowMouseCaptured || !TransparencyFieldMouseCaptured)
            {
                return;
            }

            UpdateAlpha(e.GetPosition(TransparencyField).X);
        }

        private void TransparencyField_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            TransparencyField.ReleaseMouseCapture();
            TransparencyFieldMouseCaptured = false;
            SelectedColor = GetColor();
        }

        private void ColorField_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            ColorFieldMouseCaptured = ColorField.CaptureMouse();
            Point coordinates = e.GetPosition(ColorField);
            UpdateSampleXY(coordinates.X, coordinates.Y);
        }

        private void ColorField_MouseMove(object sender, MouseEventArgs e)
        {
            if (RainbowMouseCaptured || TransparencyFieldMouseCaptured || !ColorFieldMouseCaptured)
            {
                return;
            }

            Point coordinates = e.GetPosition(ColorField);
            UpdateSampleXY(coordinates.X, coordinates.Y);
        }

        private void ColorField_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ColorField.ReleaseMouseCapture();
            ColorFieldMouseCaptured = false;
            SelectedColor = GetColor();
        }

        private Color GetColor()
        {
            double yComponent = 1 - (m_sampleY / ColorFieldHeight);
            double xComponent = m_sampleX / ColorFieldWidth;
            double hueComponent = (m_huePos / RainbowHeight) * 360;

            Color result = ColorSpace.ConvertHsvToRgb(hueComponent, xComponent, yComponent);
            result.A = (byte)(m_alpha / TransparencyFieldWidth * 255);

            return result;
        }

        private Color GetColorFromHue()
        {
            double huePos = m_huePos / RainbowHeight * 255;
            Color c = ColorSpace.GetColorFromPosition(huePos);
            return c;
        }

        private void UpdateSampleXY(double x, double y)
        {
            if (x < 0)
            {
                m_sampleX = 0;
            }
            else if (x >= ColorFieldWidth)
            {
                m_sampleX = ColorFieldWidth;
            }
            else
            {
                m_sampleX = x;
            }

            if (y < 0)
            {
                m_sampleY = 0;
            }
            else if (y >= ColorFieldHeight)
            {
                m_sampleY = ColorFieldHeight;
            }
            else
            {
                m_sampleY = y;
            }
            UpdateSelectedColorSample();
            UpdateColorSelector();
        }

        private void UpdateHuePos(double y)
        {
            if (y < 0)
            {
                m_huePos = 0;
            }
            else if (y >= RainbowHeight)
            {
                m_huePos = RainbowHeight;
            }
            else
            {
                m_huePos = y;
            }

            if (SelectedColor == Colors.Black || SelectedColor == Colors.White)
            {
                SelectedColor = GetColorFromHue();
                return;
            }

            UpdateHueSelector();
            UpdateColorFieldBackground();
            UpdateSelectedColorSample();
        }

        private void UpdateAlpha(double x)
        {
            if (x < 0)
            {
                m_alpha = 0;
            }
            else if (x >= TransparencyFieldWidth)
            {
                m_alpha = TransparencyFieldWidth;
            }
            else
            {
                m_alpha = x;
            }

            UpdateSelectedColorSample();
        }

        private void UpdateSelectedColorSample()
        {
            var color = SelectedColor;
            if (Expanded)
            {
                color = GetColor();
            }
            ColorSample.Fill = new SolidColorBrush(color);
            UpdateAlphaSelector();
            UpdateTransparency();
            FireSelectedColorChangingEvent(color);
        }

        private void UpdateColorFieldBackground()
        {
            Color c = GetColorFromHue();
            ColorField.Background = new SolidColorBrush(c);
        }

        private void UpdateTransparency()
        {
            Color c = GetColor();
            Color transparent = c;
            c.A = 255;
            transparent.A = 0;
            TransparencyGradient.Fill = new LinearGradientBrush()
            {
                GradientStops = new GradientStopCollection()
                {
                    new GradientStop() { Offset = 0, Color = transparent },
                    new GradientStop() { Offset = 1, Color = c }
                }
            };
        }

        private void UpdateHueSelector()
        {
            Canvas.SetTop(HueSelector, m_huePos - HueSelector.ActualHeight / 2);
            HueSelector.Width = Rainbow.ActualWidth;
        }

        private void UpdateAlphaSelector()
        {
            Canvas.SetLeft(AlphaSelector, m_alpha - AlphaSelector.ActualWidth / 2);
            AlphaSelector.Height = TransparencyField.ActualHeight;
        }

        private void UpdateColorSelector()
        {
            Canvas.SetLeft(ColorSelector, m_sampleX - ColorSelector.ActualWidth / 2);
            Canvas.SetTop(ColorSelector, m_sampleY - ColorSelector.ActualHeight / 2);
        }

        private double ColorFieldHeight
        {
            get
            {
                return ColorField.ActualHeight;
            }
        }

        private double ColorFieldWidth
        {
            get
            {
                return ColorField.ActualWidth;
            }
        }

        private double RainbowHeight
        {
            get
            {
                return Rainbow.ActualHeight;
            }
        }

        private double TransparencyFieldWidth
        {
            get
            {
                return TransparencyField.ActualWidth;
            }
        }

        /// <summary>
        /// Event fired when the selected color changes.  This event occurs when the 
        /// left-mouse button is lifted after clicking.
        /// </summary>
        public event SelectedColorChangedHandler SelectedColorChanged;

        /// <summary>
        /// Event fired when the selected color is changing.  This event occurs when the 
        /// left-mouse button is pressed and the user is moving the mouse.
        /// </summary>
        public event SelectedColorChangingHandler SelectedColorChanging;

        private bool RainbowMouseCaptured;
        private bool TransparencyFieldMouseCaptured;
        private bool ColorFieldMouseCaptured;
        private double m_huePos;
        private double m_alpha;
        private double m_sampleX;
        private double m_sampleY;

        #region SelectedColor Dependency Property
        /// <summary>
        /// Gets or sets the currently selected color in the Color Picker.
        /// </summary>
        public Color SelectedColor
        {
            get
            {
                return (Color)GetValue(SelectedColorProperty);
            }
            set
            {
                UpdateValuesFromColor(value);
                UpdateColorFieldBackground();
                UpdateColorSelector();
                UpdateHueSelector();
                SetValue(SelectedColorProperty, value);
                UpdateSelectedColorSample();
            }
        }

        private void UpdateValuesFromColor(Color value)
        {
            HSV hsv = ColorSpace.ConvertRgbToHsv(value);
            m_huePos = (hsv.Hue / 360 * RainbowHeight);
            m_sampleY = (1 - hsv.Value) * ColorFieldHeight;
            m_sampleX = hsv.Saturation * ColorFieldWidth;
            m_alpha = value.A / 255.0 * TransparencyFieldWidth;
        }

        /// <summary>
        /// SelectedColor Dependency Property.
        /// </summary>
        public static readonly DependencyProperty SelectedColorProperty =
            DependencyProperty.Register(
                "SelectedColor",
                typeof(Color),
                typeof(ColorPicker),
                new PropertyMetadata(Colors.Black, new PropertyChangedCallback(SelectedColorPropertyChanged)));

        private static void SelectedColorPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            ColorPicker p = d as ColorPicker;
            if (p != null && p.SelectedColorChanged != null)
            {
                SelectedColorEventArgs args = new SelectedColorEventArgs((Color)e.NewValue);
                p.SelectedColorChanged(p, args);
            }
        }

        private void FireSelectedColorChangingEvent(Color selectedColor)
        {
            if (SelectedColorChanging != null)
            {
                SelectedColorEventArgs args = new SelectedColorEventArgs(selectedColor);
                SelectedColorChanging(this, args);
            }
        }

        #endregion
    }
}
