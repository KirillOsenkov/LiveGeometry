using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Edits a brush: a solid color, or a linear gradient defined by the stops of a
/// <see cref="GradientStopBar"/> and an angle. There is a single color picker; in gradient
/// mode it shows and edits the color of the stop selected on the bar.
/// </summary>
public class BrushPickerView : Decorator
{
    const string Solid = "Solid";
    const string Gradient = "Gradient";

    readonly SegmentSwitcher kinds = new SegmentSwitcher() { Margin = new Thickness(0, 0, 0, 6) };
    readonly StackPanel gradientPanel = new StackPanel() { Margin = new Thickness(0, 0, 0, 8) };
    readonly GradientStopBar stopBar = new GradientStopBar();
    readonly Slider angleSlider = new Slider() { Minimum = 0, Maximum = 360, TickFrequency = 15, IsSnapToTickEnabled = true };
    readonly TextBlock angleText = new TextBlock() { Width = 34, TextAlignment = TextAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
    readonly ColorPickerView colorPicker;

    Color solidColor = Colors.Black;
    bool isUpdating;

    public BrushPickerView()
        : this(new ColorPickerView())
    {
    }

    /// <param name="colorPicker">The picker to use, e.g. one with a custom set of pages</param>
    public BrushPickerView(ColorPickerView colorPicker)
    {
        this.colorPicker = colorPicker;

        kinds.Add(Solid, Solid);
        kinds.Add(Gradient, Gradient);
        kinds.Current = Solid;

        var angleRow = new DockPanel() { Margin = new Thickness(0, 2, 0, 0) };
        var angleLabel = new TextBlock() { Text = "Angle", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 8, 0) };
        DockPanel.SetDock(angleLabel, Dock.Left);
        DockPanel.SetDock(angleText, Dock.Right);
        angleRow.Children.Add(angleLabel);
        angleRow.Children.Add(angleText);
        angleRow.Children.Add(angleSlider);

        gradientPanel.Children.Add(stopBar);
        gradientPanel.Children.Add(angleRow);
        gradientPanel.IsVisible = false;

        var layout = new StackPanel() { Width = ColorPickerView.ContentWidth, HorizontalAlignment = HorizontalAlignment.Left };
        layout.Children.Add(kinds);
        layout.Children.Add(gradientPanel);
        layout.Children.Add(colorPicker);
        Child = layout;

        kinds.Selected += kind => SwitchKind((string)kind);
        stopBar.SelectionChanged += () =>
        {
            if (IsGradient && stopBar.SelectedStop != null)
            {
                colorPicker.Color = stopBar.SelectedStop.Color;
            }
        };
        stopBar.StopsChanged += RaiseBrushChanged;
        angleSlider.PropertyChanged += (s, e) =>
        {
            if (e.Property == Slider.ValueProperty)
            {
                angleText.Text = ((int)angleSlider.Value) + "°";
                RaiseBrushChanged();
            }
        };
        colorPicker.ColorChanged += picked =>
        {
            if (IsGradient)
            {
                stopBar.SetSelectedColor(picked);
            }
            else
            {
                solidColor = picked;
                RaiseBrushChanged();
            }
        };

        angleText.Text = "0°";
    }

    /// <summary>Turn off to offer solid colors only.</summary>
    public bool EnableGradients
    {
        get => kinds.IsVisible;
        set => kinds.IsVisible = value;
    }

    bool IsGradient => (string)kinds.Current == Gradient;

    /// <summary>The user changed something.</summary>
    public event Action<IBrush> BrushChanged;

    /// <summary>
    /// The brush as currently edited; a new instance every time. Setting it does not raise
    /// <see cref="BrushChanged"/>. Anything other than a solid or linear gradient brush
    /// shows up as black.
    /// </summary>
    public IBrush Brush
    {
        get
        {
            if (!IsGradient)
            {
                return new SolidColorBrush(solidColor);
            }

            double radians = angleSlider.Value * System.Math.PI / 180;
            double x = System.Math.Cos(radians) / 2;
            double y = System.Math.Sin(radians) / 2;
            var brush = new LinearGradientBrush()
            {
                StartPoint = new RelativePoint(0.5 - x, 0.5 - y, RelativeUnit.Relative),
                EndPoint = new RelativePoint(0.5 + x, 0.5 + y, RelativeUnit.Relative)
            };
            foreach (var stop in stopBar.Stops)
            {
                brush.GradientStops.Add(new GradientStop(stop.Color, stop.Offset));
            }

            return brush;
        }

        set
        {
            isUpdating = true;
            try
            {
                if (value is ILinearGradientBrush linear && linear.GradientStops.Count >= 2)
                {
                    var direction = linear.EndPoint.Point - linear.StartPoint.Point;
                    double degrees = System.Math.Atan2(direction.Y, direction.X) * 180 / System.Math.PI;
                    angleSlider.Value = System.Math.Round((degrees + 360) % 360);
                    stopBar.SetStops(linear.GradientStops.Select(s => new ColorStop(s.Offset, s.Color)));
                    ShowKind(Gradient);
                }
                else
                {
                    solidColor = (value as ISolidColorBrush)?.Color ?? Colors.Black;
                    ShowKind(Solid);
                }
            }
            finally
            {
                isUpdating = false;
            }
        }
    }

    void SwitchKind(string kind)
    {
        if (kind == Gradient && stopBar.Stops.Count < 2)
        {
            // start from the solid color, fading to white
            stopBar.SetStops(new[] { new ColorStop(0, solidColor), new ColorStop(1, Colors.White) });
        }
        else if (kind == Solid && stopBar.SelectedStop != null)
        {
            solidColor = stopBar.SelectedStop.Color;
        }

        ShowKind(kind);
        RaiseBrushChanged();
    }

    void ShowKind(string kind)
    {
        kinds.Current = kind;
        gradientPanel.IsVisible = kind == Gradient;
        colorPicker.Color = kind == Gradient && stopBar.SelectedStop != null
            ? stopBar.SelectedStop.Color
            : solidColor;
    }

    void RaiseBrushChanged()
    {
        if (!isUpdating)
        {
            BrushChanged?.Invoke(Brush);
        }
    }
}
