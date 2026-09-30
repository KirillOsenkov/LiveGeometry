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
    readonly Avalonia.Controls.Slider angleSlider = new Avalonia.Controls.Slider() { Minimum = 0, Maximum = 360, TickFrequency = 15, IsSnapToTickEnabled = true };
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
            if (e.Property == Avalonia.Controls.Slider.ValueProperty)
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

    /// <summary>The theme color of the background the picker sits on, so that the tabs can blend into it.</summary>
    public string Surface
    {
        get => kinds.Surface;
        set
        {
            kinds.Surface = value;
            colorPicker.Surface = value;
        }
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

            // The line goes through the middle of the box at the angle, and is as long as it
            // takes for the first and the last color to be reached in the corners (as CSS
            // does it): (0,0) to (1,1) at 45 degrees, edge to edge at 0 and 90. The stops
            // then reach everywhere. A shorter line leaves the corners flat, where no stop
            // can do anything; this one ends outside the box at angles in between.
            double radians = angleSlider.Value * System.Math.PI / 180;
            double cosine = System.Math.Cos(radians);
            double sine = System.Math.Sin(radians);
            double reach = (System.Math.Abs(cosine) + System.Math.Abs(sine)) / 2;
            double x = cosine * reach;
            double y = sine * reach;
            var brush = new LinearGradientBrush()
            {
                StartPoint = new RelativePoint(Tidy(0.5 - x), Tidy(0.5 - y), RelativeUnit.Relative),
                EndPoint = new RelativePoint(Tidy(0.5 + x), Tidy(0.5 + y), RelativeUnit.Relative)
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

    /// <summary>Without the rounding error of the sine and cosine: 0 and 1 where they are meant, and no negative zero</summary>
    static double Tidy(double coordinate)
    {
        return System.Math.Round(coordinate, digits: 6) + 0.0;
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
